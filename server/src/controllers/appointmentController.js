const Appointment = require('../models/Appointment');
const Patient = require('../models/Patient');
const Doctor = require('../models/Doctor');
const neo4jService = require('../services/neo4jService');

// CREATE Appointment
exports.createAppointment = async (req, res, next) => {
  try {
    const aptData = req.body;
    aptData.appointmentId = aptData.appointmentId.toUpperCase().trim();
    aptData.patientId = aptData.patientId.toUpperCase().trim();
    aptData.doctorId = aptData.doctorId.toUpperCase().trim();

    // Check duplicate ID
    const existing = await Appointment.findOne({ appointmentId: aptData.appointmentId });
    if (existing) {
      return res.status(409).json({
        success: false,
        error: `Appointment with ID ${aptData.appointmentId} already exists.`,
      });
    }

    // Verify patient and doctor exist
    const patientExists = await Patient.findOne({ patientId: aptData.patientId });
    if (!patientExists) {
      return res.status(404).json({
        success: false,
        error: `Referenced Patient ${aptData.patientId} does not exist in database.`,
      });
    }

    const doctorExists = await Doctor.findOne({ doctorId: aptData.doctorId });
    if (!doctorExists) {
      return res.status(404).json({
        success: false,
        error: `Referenced Doctor ${aptData.doctorId} does not exist in database.`,
      });
    }

    const appointment = await Appointment.create(aptData);

    // Create TREATED_BY relationship in Neo4j
    await neo4jService.createRelationship({
      fromLabel: 'Patient',
      fromIdKey: 'patientId',
      fromIdValue: appointment.patientId,
      relType: 'TREATED_BY',
      toLabel: 'Doctor',
      toIdKey: 'doctorId',
      toIdValue: appointment.doctorId,
    });

    res.status(201).json({
      success: true,
      message: 'Appointment booked successfully and graph relationship created',
      data: appointment,
    });
  } catch (error) {
    next(error);
  }
};

// READ ALL Appointments
exports.getAllAppointments = async (req, res, next) => {
  try {
    const { patientId, doctorId, status, date } = req.query;
    const filter = {};

    if (patientId) filter.patientId = patientId.toUpperCase();
    if (doctorId) filter.doctorId = doctorId.toUpperCase();
    if (status) filter.status = status;
    if (date) filter.date = date;

    const appointments = await Appointment.find(filter).sort({ date: -1, time: -1 });
    res.status(200).json({
      success: true,
      count: appointments.length,
      data: appointments,
    });
  } catch (error) {
    next(error);
  }
};

// READ Appointment by ID
exports.getAppointmentById = async (req, res, next) => {
  try {
    const { id } = req.params;
    const appointment = await Appointment.findOne({ appointmentId: id.toUpperCase().trim() });

    if (!appointment) {
      return res.status(404).json({
        success: false,
        error: `Appointment with ID '${id}' not found`,
      });
    }

    res.status(200).json({
      success: true,
      data: appointment,
    });
  } catch (error) {
    next(error);
  }
};

// UPDATE Appointment
exports.updateAppointment = async (req, res, next) => {
  try {
    const { id } = req.params;
    const updates = { ...req.body };
    delete updates.appointmentId;

    if (updates.patientId) updates.patientId = updates.patientId.toUpperCase().trim();
    if (updates.doctorId) updates.doctorId = updates.doctorId.toUpperCase().trim();

    const appointment = await Appointment.findOneAndUpdate(
      { appointmentId: id.toUpperCase().trim() },
      { $set: updates },
      { new: true, runValidators: true }
    );

    if (!appointment) {
      return res.status(404).json({
        success: false,
        error: `Appointment with ID '${id}' not found`,
      });
    }

    res.status(200).json({
      success: true,
      message: 'Appointment updated successfully',
      data: appointment,
    });
  } catch (error) {
    next(error);
  }
};

// DELETE Appointment
exports.deleteAppointment = async (req, res, next) => {
  try {
    const { id } = req.params;
    const appointment = await Appointment.findOneAndDelete({ appointmentId: id.toUpperCase().trim() });

    if (!appointment) {
      return res.status(404).json({
        success: false,
        error: `Appointment with ID '${id}' not found`,
      });
    }

    res.status(200).json({
      success: true,
      message: `Appointment ${id} deleted successfully`,
      data: appointment,
    });
  } catch (error) {
    next(error);
  }
};
