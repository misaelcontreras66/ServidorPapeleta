// server.js
//import express, {static, json } from 'express';
import express from 'express';
import exceljs from 'exceljs';
import os from 'os';
import fs from 'fs';
import { execSync } from 'child_process';

function getLocalIP() {
  const interfaces = os.networkInterfaces();

  for (let name in interfaces) {
    for (let net of interfaces[name]) {
      if (net.family === 'IPv4' && !net.internal) {
        return net.address;
      }
    }
  }
}

function updateGitHub(ip, port) {
  const data = {
    ip,
    puerto: port
  };

  fs.writeFileSync('./ip.json', JSON.stringify(data, null, 2));

  try {
    console.log("1️⃣ git add...");
    execSync('git add ip.json', { stdio: 'inherit' });

    console.log("2️⃣ git commit...");
    execSync(`git commit -m "update ip ${ip}"`, { stdio: 'inherit' });

    console.log("3️⃣ git push...");
    execSync('git push origin main', { stdio: 'inherit' });

    console.log("✅ GitHub actualizado");
  } catch (err) {
    console.error("❌ Falló un comando de Git.");
  }
}

const app = express();
const port = 3000;

let solicitudes = [];

app.use(express.static('public'));
app.use(express.json());
app.get('/get-requests', (req, res) => res.json(solicitudes));

app.post('/submit-request', (req, res) => {
  const nuevaSolicitud = {
    ...req.body,
    fecha: new Date().toLocaleString('es-ES', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    })
  };
  solicitudes.push(nuevaSolicitud);
  res.status(201).send({ status: 'success' });
});

app.put('/update-request/:id', (req, res) => {
  const index = parseInt(req.params.id);
  if (index >= 0 && index < solicitudes.length) {
    solicitudes[index] = {
      ...req.body,
      fecha: new Date().toLocaleString('es-ES', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      })
    };
    res.status(200).send({ status: 'updated' });
  } else {
    res.status(404).send({ error: 'Solicitud no encontrada' });
  }
});

app.delete('/delete-request/:id', (req, res) => {
  const index = parseInt(req.params.id);
  if (index >= 0 && index < solicitudes.length) {
    solicitudes.splice(index, 1);
    res.status(200).send({ status: 'deleted' });
  } else {
    res.status(404).send({ error: 'Solicitud no encontrada' });
  }
});

app.get('/download-excel', async (req, res) => {
  try {
    const workbook = new exceljs.Workbook();
    const worksheet = workbook.addWorksheet('Solicitudes');


    worksheet.columns = [
      { header: 'Estudiante', key: 'estudiante', width: 25 },
      { header: 'Matrícula', key: 'matricula', width: 15 },
      { header: 'Profesor', key: 'profesor', width: 30 },
      { header: 'UEA', key: 'uea', width: 15 },
      { header: 'Tipo Práctica', key: 'tipoPractica', width: 20 },
      { header: 'Materiales', key: 'materiales', width: 50 },
      { header: 'Fecha', key: 'fecha', width: 20 }
    ];

    solicitudes.forEach(sol => {
      worksheet.addRow({
        ...sol,
        materiales: sol.materiales.map(m => `${m.nombre} (${m.cantidad})`).join('\n')
      });
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=solicitudes.xlsx');
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    res.status(500).send('Error al generar Excel');
  }
});

app.listen(port, '0.0.0.0', () => {
  console.log('══════════════════════════════════════');
  console.log(` Servidor activo en:`);
  console.log(`• Local:    http://localhost:${port}`);
  const ip = getLocalIP();
  updateGitHub(ip, port);
  console.log(`• Red:      http://${ip}:${port}`);
  console.log('══════════════════════════════════════');
});