const Prescription = require('../models/Prescription');
const Patient = require('../models/Patient');
const Doctor = require('../models/Doctor');
const neo4jService = require('../services/neo4jService');

// CREATE Prescription
exports.createPrescription = async (req, res, next) => {
  try {
    const rxData = req.body;
    rxData.prescriptionId = rxData.prescriptionId.toUpperCase().trim();
    rxData.patientId = rxData.patientId.toUpperCase().trim();
    rxData.doctorId = rxData.doctorId.toUpperCase().trim();

    // Check duplicate
    const existing = await Prescription.findOne({ prescriptionId: rxData.prescriptionId });
    if (existing) {
      return res.status(409).json({
        success: false,
        error: `Prescription with ID ${rxData.prescriptionId} already exists.`,
      });
    }

    const patient = await Patient.findOne({ patientId: rxData.patientId });
    if (!patient) {
      return res.status(404).json({
        success: false,
        error: `Referenced Patient ${rxData.patientId} does not exist.`,
      });
    }

    const doctor = await Doctor.findOne({ doctorId: rxData.doctorId });
    if (!doctor) {
      return res.status(404).json({
        success: false,
        error: `Referenced Doctor ${rxData.doctorId} does not exist.`,
      });
    }

    const prescription = await Prescription.create(rxData);

    // Synchronize relationships to Neo4j graph:
    // 1. Patient - TREATED_BY -> Doctor
    await neo4jService.createRelationship({
      fromLabel: 'Patient',
      fromIdKey: 'patientId',
      fromIdValue: prescription.patientId,
      relType: 'TREATED_BY',
      toLabel: 'Doctor',
      toIdKey: 'doctorId',
      toIdValue: prescription.doctorId,
    });

    // 2. Doctor - PRESCRIBED -> Medicine and Patient - TAKES -> Medicine
    if (Array.isArray(prescription.medicines)) {
      for (const med of prescription.medicines) {
        if (med.medicineId) {
          // Doctor PRESCRIBED Medicine
          await neo4jService.createRelationship({
            fromLabel: 'Doctor',
            fromIdKey: 'doctorId',
            fromIdValue: prescription.doctorId,
            relType: 'PRESCRIBED',
            toLabel: 'Medicine',
            toIdKey: 'medicineId',
            toIdValue: med.medicineId.toUpperCase(),
          });

          // Patient TAKES Medicine
          await neo4jService.createRelationship({
            fromLabel: 'Patient',
            fromIdKey: 'patientId',
            fromIdValue: prescription.patientId,
            relType: 'TAKES',
            toLabel: 'Medicine',
            toIdKey: 'medicineId',
            toIdValue: med.medicineId.toUpperCase(),
          });
        }
      }
    }

    res.status(201).json({
      success: true,
      message: 'Prescription recorded in MongoDB and graph edges synced in Neo4j',
      data: prescription,
    });
  } catch (error) {
    next(error);
  }
};

// READ ALL Prescriptions
exports.getAllPrescriptions = async (req, res, next) => {
  try {
    const { patientId, doctorId } = req.query;
    const filter = {};

    if (patientId) filter.patientId = patientId.toUpperCase();
    if (doctorId) filter.doctorId = doctorId.toUpperCase();

    const prescriptions = await Prescription.find(filter).sort({ date: -1, createdAt: -1 });
    res.status(200).json({
      success: true,
      count: prescriptions.length,
      data: prescriptions,
    });
  } catch (error) {
    next(error);
  }
};

// READ Prescription by ID
exports.getPrescriptionById = async (req, res, next) => {
  try {
    const { id } = req.params;
    const prescription = await Prescription.findOne({ prescriptionId: id.toUpperCase().trim() });

    if (!prescription) {
      return res.status(404).json({
        success: false,
        error: `Prescription with ID '${id}' not found`,
      });
    }

    res.status(200).json({
      success: true,
      data: prescription,
    });
  } catch (error) {
    next(error);
  }
};

// DELETE Prescription
exports.deletePrescription = async (req, res, next) => {
  try {
    const { id } = req.params;
    const prescription = await Prescription.findOneAndDelete({ prescriptionId: id.toUpperCase().trim() });

    if (!prescription) {
      return res.status(404).json({
        success: false,
        error: `Prescription with ID '${id}' not found`,
      });
    }

    res.status(200).json({
      success: true,
      message: `Prescription ${id} deleted successfully`,
      data: prescription,
    });
  } catch (error) {
    next(error);
  }
};
