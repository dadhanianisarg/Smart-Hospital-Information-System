import React, { useEffect, useState } from 'react';
import { api } from '../services/api';
import { Plus, Trash2, Pill, X } from 'lucide-react';

export default function Prescriptions() {
  const [prescriptions, setPrescriptions] = useState([]);
  const [patients, setPatients] = useState([]);
  const [doctors, setDoctors] = useState([]);
  const [medicines, setMedicines] = useState([]);
  const [loading, setLoading] = useState(true);

  const [showAddModal, setShowAddModal] = useState(false);
  const initialForm = {
    prescriptionId: '',
    patientId: '',
    doctorId: '',
    diagnosis: '',
    instructions: '',
    date: new Date().toISOString().split('T')[0],
    medicineId: 'MED01',
    dosage: '5mg',
    frequency: 'Once daily',
    duration: '30 days',
  };
  const [formData, setFormData] = useState(initialForm);

  const fetchData = async () => {
    setLoading(true);
    try {
      const [rxRes, patRes, docRes, medRes] = await Promise.all([
        api.getPrescriptions(),
        api.getPatients(),
        api.getDoctors(),
        api.getMedicines(),
      ]);
      if (rxRes.success) setPrescriptions(rxRes.data);
      if (patRes.success) {
        setPatients(patRes.data);
        if (patRes.data.length > 0 && !formData.patientId) {
          setFormData(f => ({ ...f, patientId: patRes.data[0].patientId }));
        }
      }
      if (docRes.success) {
        setDoctors(docRes.data);
        if (docRes.data.length > 0 && !formData.doctorId) {
          setFormData(f => ({ ...f, doctorId: docRes.data[0].doctorId }));
        }
      }
      if (medRes.success) {
        setMedicines(medRes.data);
      }
    } catch (err) {
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
  }, []);

  const handleCreate = async (e) => {
    e.preventDefault();
    try {
      const selectedMed = medicines.find(m => m.medicineId === formData.medicineId);
      const payload = {
        prescriptionId: formData.prescriptionId,
        patientId: formData.patientId,
        doctorId: formData.doctorId,
        diagnosis: formData.diagnosis,
        instructions: formData.instructions,
        date: formData.date,
        medicines: [
          {
            medicineId: formData.medicineId,
            name: selectedMed ? selectedMed.name : formData.medicineId,
            dosage: formData.dosage,
            frequency: formData.frequency,
            duration: formData.duration,
          },
        ],
      };
      await api.createPrescription(payload);
      setShowAddModal(false);
      setFormData(initialForm);
      fetchData();
    } catch (err) {
      alert(`Error creating prescription: ${err.message}`);
    }
  };

  const handleDelete = async (prescriptionId) => {
    if (window.confirm(`Delete prescription ${prescriptionId}?`)) {
      try {
        await api.deletePrescription(prescriptionId);
        fetchData();
      } catch (err) {
        alert(`Error deleting prescription: ${err.message}`);
      }
    }
  };

  return (
    <div>
      <div className="card">
        <div className="card-header">
          <div className="card-title">Prescriptions (Embedded Medicines Modeling)</div>
          <button className="btn btn-primary" onClick={() => setShowAddModal(true)}>
            <Plus size={16} /> New Prescription
          </button>
        </div>

        <div className="table-responsive">
          <table>
            <thead>
              <tr>
                <th>Rx ID</th>
                <th>Patient ID</th>
                <th>Doctor ID</th>
                <th>Diagnosis</th>
                <th>Embedded Medicines</th>
                <th>Date</th>
                <th>Instructions</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              {loading ? (
                <tr><td colSpan="8" style={{ textAlign: 'center' }}>Loading prescriptions...</td></tr>
              ) : prescriptions.length === 0 ? (
                <tr><td colSpan="8" style={{ textAlign: 'center', color: '#94a3b8' }}>No prescriptions found</td></tr>
              ) : (
                prescriptions.map((rx) => (
                  <tr key={rx.prescriptionId}>
                    <td><strong>{rx.prescriptionId}</strong></td>
                    <td><span className="badge badge-blue">{rx.patientId}</span></td>
                    <td><span className="badge" style={{ backgroundColor: '#dcfce7', color: '#166534' }}>{rx.doctorId}</span></td>
                    <td><strong>{rx.diagnosis}</strong></td>
                    <td>
                      {rx.medicines && rx.medicines.map((m, idx) => (
                        <div key={idx} style={{ fontSize: '0.8rem', color: '#475569', marginBottom: '2px' }}>
                          💊 <strong>{m.name}</strong> ({m.dosage} - {m.frequency}, {m.duration})
                        </div>
                      ))}
                    </td>
                    <td>{rx.date}</td>
                    <td style={{ fontSize: '0.82rem', color: '#64748b' }}>{rx.instructions || 'N/A'}</td>
                    <td>
                      <button className="btn btn-danger btn-sm" onClick={() => handleDelete(rx.prescriptionId)}>
                        <Trash2 size={14} />
                      </button>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* CREATE MODAL */}
      {showAddModal && (
        <div className="modal-overlay">
          <div className="modal-content">
            <div className="modal-header">
              <h3>Create Prescription Document</h3>
              <button onClick={() => setShowAddModal(false)} style={{ background: 'none', border: 'none', cursor: 'pointer' }}><X size={20} /></button>
            </div>
            <form onSubmit={handleCreate}>
              <div className="form-group">
                <label>Prescription ID *</label>
                <input type="text" className="form-control" required placeholder="e.g. RX109" value={formData.prescriptionId} onChange={(e) => setFormData({ ...formData, prescriptionId: e.target.value })} />
              </div>
              <div className="form-row">
                <div className="form-group">
                  <label>Patient *</label>
                  <select className="form-control" value={formData.patientId} onChange={(e) => setFormData({ ...formData, patientId: e.target.value })}>
                    {patients.map(p => (
                      <option key={p.patientId} value={p.patientId}>{p.name} ({p.patientId})</option>
                    ))}
                  </select>
                </div>
                <div className="form-group">
                  <label>Doctor *</label>
                  <select className="form-control" value={formData.doctorId} onChange={(e) => setFormData({ ...formData, doctorId: e.target.value })}>
                    {doctors.map(d => (
                      <option key={d.doctorId} value={d.doctorId}>{d.name} ({d.specialization})</option>
                    ))}
                  </select>
                </div>
              </div>
              <div className="form-row">
                <div className="form-group">
                  <label>Diagnosis *</label>
                  <input type="text" className="form-control" required placeholder="e.g. Hypertension" value={formData.diagnosis} onChange={(e) => setFormData({ ...formData, diagnosis: e.target.value })} />
                </div>
                <div className="form-group">
                  <label>Prescription Date *</label>
                  <input type="date" className="form-control" required value={formData.date} onChange={(e) => setFormData({ ...formData, date: e.target.value })} />
                </div>
              </div>

              <div style={{ borderTop: '1px solid #e2e8f0', paddingTop: '12px', marginTop: '12px', marginBottom: '12px' }}>
                <h4 style={{ fontSize: '0.9rem', marginBottom: '8px', color: '#0f172a' }}>Embedded Medicine Information</h4>
                <div className="form-row">
                  <div className="form-group">
                    <label>Medicine *</label>
                    <select className="form-control" value={formData.medicineId} onChange={(e) => setFormData({ ...formData, medicineId: e.target.value })}>
                      {medicines.map(m => (
                        <option key={m.medicineId} value={m.medicineId}>{m.name} ({m.category})</option>
                      ))}
                    </select>
                  </div>
                  <div className="form-group">
                    <label>Dosage</label>
                    <input type="text" className="form-control" value={formData.dosage} onChange={(e) => setFormData({ ...formData, dosage: e.target.value })} />
                  </div>
                </div>
                <div className="form-row">
                  <div className="form-group">
                    <label>Frequency</label>
                    <input type="text" className="form-control" value={formData.frequency} onChange={(e) => setFormData({ ...formData, frequency: e.target.value })} />
                  </div>
                  <div className="form-group">
                    <label>Duration</label>
                    <input type="text" className="form-control" value={formData.duration} onChange={(e) => setFormData({ ...formData, duration: e.target.value })} />
                  </div>
                </div>
              </div>

              <div className="form-group">
                <label>Instructions / Advice</label>
                <input type="text" className="form-control" placeholder="e.g. Take after food with water" value={formData.instructions} onChange={(e) => setFormData({ ...formData, instructions: e.target.value })} />
              </div>

              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '12px', marginTop: '20px' }}>
                <button type="button" className="btn btn-secondary" onClick={() => setShowAddModal(false)}>Cancel</button>
                <button type="submit" className="btn btn-primary">Save Prescription</button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
