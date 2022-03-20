const excelJs = require('exceljs');

const {
  initializeApp,
  applicationDefault,
  cert,
} = require('firebase-admin/app');
const {
  getFirestore,
  Timestamp,
  FieldValue,
} = require('firebase-admin/firestore');

const serviceAccount = require('./serviceAccountKey.json');

initializeApp({
  credential: cert(serviceAccount),
});

const db = getFirestore();

const stock = [];

const getData = async () => {
  const snapshot = await db.collection('products').get();
  snapshot.forEach((doc) => {
    const id = doc.id;
    const warehouse = 'KHH-HCM';
    let quantity = 0;
    doc.data().inventory.forEach((count) => {
      switch (count.unit) {
        case 'đôi':
          quantity += 1 * count.quantity;
          break;
        case 'bao 120':
          quantity += 120 * count.quantity;
          break;
        case 'bao 60':
          quantity += 60 * count.quantity;
          break;
        case 'bịch 6':
          quantity += 6 * count.quantity;
          break;
        case 'thùng':
          quantity += 6 * count.quantity;
          break;
        case 'bịch 12':
          quantity += 12 * count.quantity;
          break;
      }
    });
    const quality = quantity;
    stock.push({ id, warehouse, quantity, quality });
  });
};

const printData = async () => {
  await getData();

  const workbook = new excelJs.Workbook();
  const worksheet = workbook.addWorksheet('Kiểm kê vật tư hàng hóa');
  const path = './';
  worksheet.columns = [
    { header: 'Mã hàng (*)', key: 'id', width: 10 },
    { header: 'Kho (*)', key: 'warehouse', width: 10 },
    { header: 'Số lượng theo kiểm kê', key: 'quantity', width: 10 },
    { header: 'Còn tốt 100%', key: 'quality', width: 10 },
  ];

  stock.forEach((data) => {
    worksheet.addRow(data);
  });

  try {
    const data = await workbook.xlsx
      .writeFile(`${path}/Kiem_ke_vat_tu_hang_hoa.xlsx`)
      .then(() => console.log('file successfully generated'));
  } catch (err) {
    console.log('Something went wrong');
  }
};

printData();
