const Doctor = require('../models/Doctor');
const neo4jService = require('../services/neo4jService');

// CREATE Doctor
exports.createDoctor = async (req, res, next) => {
  try {
    const doctorData = req.body;
    doctorData.doctorId = doctorData.doctorId.toUpperCase().trim();
    doctorData.departmentId = doctorData.departmentId.toUpperCase().trim();

    const existing = await Doctor.findOne({ doctorId: doctorData.doctorId });
    if (existing) {
      return res.status(409).json({
        success: false,
        error: `Doctor with ID ${doctorData.doctorId} already exists.`,
      });
    }

    const doctor = await Doctor.create(doctorData);

    // Sync to Neo4j
    await neo4jService.createDoctorNode({
      doctorId: doctor.doctorId,
      name: doctor.name,
      specialization: doctor.specialization,
    });

    // Relate to Department in Neo4j
    await neo4jService.createRelationship({
      fromLabel: 'Doctor',
      fromIdKey: 'doctorId',
      fromIdValue: doctor.doctorId,
      relType: 'BELONGS_TO',
      toLabel: 'Department',
      toIdKey: 'departmentId',
      toIdValue: doctor.departmentId,
    });

    res.status(201).json({
      success: true,
      message: 'Doctor created successfully in MongoDB and linked in Neo4j',
      data: doctor,
    });
  } catch (error) {
    next(error);
  }
};

// READ ALL Doctors
exports.getAllDoctors = async (req, res, next) => {
  try {
    const { departmentId, search } = req.query;
    const filter = {};

    if (departmentId) filter.departmentId = departmentId.toUpperCase();
    if (search) {
      filter.$or = [
        { name: { $regex: search, $options: 'i' } },
        { specialization: { $regex: search, $options: 'i' } },
        { doctorId: { $regex: search, $options: 'i' } },
      ];
    }

    const doctors = await Doctor.find(filter).sort({ name: 1 });
    res.status(200).json({
      success: true,
      count: doctors.length,
      data: doctors,
    });
  } catch (error) {
    next(error);
  }
};

// READ Doctor by ID
exports.getDoctorById = async (req, res, next) => {
  try {
    const { id } = req.params;
    const doctor = await Doctor.findOne({ doctorId: id.toUpperCase().trim() });

    if (!doctor) {
      return res.status(404).json({
        success: false,
        error: `Doctor with ID '${id}' not found`,
      });
    }

    res.status(200).json({
      success: true,
      data: doctor,
    });
  } catch (error) {
    next(error);
  }
};

// UPDATE Doctor
exports.updateDoctor = async (req, res, next) => {
  try {
    const { id } = req.params;
    const updates = { ...req.body };
    delete updates.doctorId;

    if (updates.departmentId) {
      updates.departmentId = updates.departmentId.toUpperCase().trim();
    }

    const doctor = await Doctor.findOneAndUpdate(
      { doctorId: id.toUpperCase().trim() },
      { $set: updates },
      { new: true, runValidators: true }
    );

    if (!doctor) {
      return res.status(404).json({
        success: false,
        error: `Doctor with ID '${id}' not found`,
      });
    }

    // Sync to Neo4j
    await neo4jService.createDoctorNode({
      doctorId: doctor.doctorId,
      name: doctor.name,
      specialization: doctor.specialization,
    });

    if (doctor.departmentId) {
      await neo4jService.createRelationship({
        fromLabel: 'Doctor',
        fromIdKey: 'doctorId',
        fromIdValue: doctor.doctorId,
        relType: 'BELONGS_TO',
        toLabel: 'Department',
        toIdKey: 'departmentId',
        toIdValue: doctor.departmentId,
      });
    }

    res.status(200).json({
      success: true,
      message: 'Doctor updated successfully',
      data: doctor,
    });
  } catch (error) {
    next(error);
  }
};

// DELETE Doctor
exports.deleteDoctor = async (req, res, next) => {
  try {
    const { id } = req.params;
    const doctor = await Doctor.findOneAndDelete({ doctorId: id.toUpperCase().trim() });

    if (!doctor) {
      return res.status(404).json({
        success: false,
        error: `Doctor with ID '${id}' not found`,
      });
    }

    res.status(200).json({
      success: true,
      message: `Doctor ${id} deleted successfully from MongoDB`,
      data: doctor,
    });
  } catch (error) {
    next(error);
  }
};
