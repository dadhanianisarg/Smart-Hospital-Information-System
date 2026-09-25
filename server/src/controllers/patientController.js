const Patient = require('../models/Patient');
const neo4jService = require('../services/neo4jService');

// CREATE Patient
exports.createPatient = async (req, res, next) => {
  try {
    const patientData = req.body;
    patientData.patientId = patientData.patientId.toUpperCase().trim();

    // Check duplicate
    const existing = await Patient.findOne({ patientId: patientData.patientId });
    if (existing) {
      return res.status(409).json({
        success: false,
        error: `Patient with ID ${patientData.patientId} already exists.`,
      });
    }

    const patient = await Patient.create(patientData);

    // Sync to Neo4j Graph
    await neo4jService.createPatientNode({
      patientId: patient.patientId,
      name: patient.name,
      age: patient.age,
      gender: patient.gender,
    });

    res.status(201).json({
      success: true,
      message: 'Patient registered successfully in MongoDB and Neo4j graph',
      data: patient,
    });
  } catch (error) {
    next(error);
  }
};

// READ ALL Patients
exports.getAllPatients = async (req, res, next) => {
  try {
    const { search, gender, bloodGroup } = req.query;
    const filter = {};

    if (gender) filter.gender = gender;
    if (bloodGroup) filter.bloodGroup = bloodGroup;
    if (search) {
      filter.$or = [
        { name: { $regex: search, $options: 'i' } },
        { patientId: { $regex: search, $options: 'i' } },
        { phone: { $regex: search, $options: 'i' } },
      ];
    }

    const patients = await Patient.find(filter).sort({ createdAt: -1 });
    res.status(200).json({
      success: true,
      count: patients.length,
      data: patients,
    });
  } catch (error) {
    next(error);
  }
};

// READ Patient by ID (patientId)
exports.getPatientById = async (req, res, next) => {
  try {
    const { id } = req.params;
    const patient = await Patient.findOne({ patientId: id.toUpperCase().trim() });

    if (!patient) {
      return res.status(404).json({
        success: false,
        error: `Patient with ID '${id}' not found`,
      });
    }

    res.status(200).json({
      success: true,
      data: patient,
    });
  } catch (error) {
    next(error);
  }
};

// UPDATE Patient
exports.updatePatient = async (req, res, next) => {
  try {
    const { id } = req.params;
    const updates = { ...req.body };
    delete updates.patientId; // Immutable identifier

    const patient = await Patient.findOneAndUpdate(
      { patientId: id.toUpperCase().trim() },
      { $set: updates },
      { new: true, runValidators: true }
    );

    if (!patient) {
      return res.status(404).json({
        success: false,
        error: `Patient with ID '${id}' not found`,
      });
    }

    // Sync to Neo4j
    await neo4jService.createPatientNode({
      patientId: patient.patientId,
      name: patient.name,
      age: patient.age,
      gender: patient.gender,
    });

    res.status(200).json({
      success: true,
      message: 'Patient updated successfully in MongoDB and Neo4j',
      data: patient,
    });
  } catch (error) {
    next(error);
  }
};

// DELETE Patient
exports.deletePatient = async (req, res, next) => {
  try {
    const { id } = req.params;
    const patient = await Patient.findOneAndDelete({ patientId: id.toUpperCase().trim() });

    if (!patient) {
      return res.status(404).json({
        success: false,
        error: `Patient with ID '${id}' not found`,
      });
    }

    res.status(200).json({
      success: true,
      message: `Patient ${id} deleted successfully from MongoDB`,
      data: patient,
    });
  } catch (error) {
    next(error);
  }
};
