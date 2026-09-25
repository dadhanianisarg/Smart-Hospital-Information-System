const Medicine = require('../models/Medicine');
const neo4jService = require('../services/neo4jService');

exports.createMedicine = async (req, res, next) => {
  try {
    const data = req.body;
    data.medicineId = data.medicineId.toUpperCase().trim();

    const existing = await Medicine.findOne({ medicineId: data.medicineId });
    if (existing) {
      return res.status(409).json({ success: false, error: `Medicine ${data.medicineId} already exists.` });
    }

    const medicine = await Medicine.create(data);
    await neo4jService.createMedicineNode({
      medicineId: medicine.medicineId,
      name: medicine.name,
      category: medicine.category,
    });

    res.status(201).json({ success: true, data: medicine });
  } catch (error) {
    next(error);
  }
};

exports.getAllMedicines = async (req, res, next) => {
  try {
    const medicines = await Medicine.find().sort({ name: 1 });
    res.status(200).json({ success: true, count: medicines.length, data: medicines });
  } catch (error) {
    next(error);
  }
};

exports.getMedicineById = async (req, res, next) => {
  try {
    const medicine = await Medicine.findOne({ medicineId: req.params.id.toUpperCase().trim() });
    if (!medicine) return res.status(404).json({ success: false, error: 'Medicine not found' });
    res.status(200).json({ success: true, data: medicine });
  } catch (error) {
    next(error);
  }
};
