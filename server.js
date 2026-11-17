
const express = require("express");
const cors = require("cors");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const app = express();
const PORT = 3000;

// Exact Excel file location
const EXCEL_FILE = path.join(__dirname, "orders.xlsx");

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, "../frontend")));

// Prevent simultaneous Excel read/write operations
let saveQueue = Promise.resolve();

function queueSave(task) {
    const current = saveQueue.then(task, task);
    saveQueue = current.catch(() => {});
    return current;
}

// Retry when Excel file is temporarily locked
async function saveWorkbook(workbook, retries = 5) {
    for (let attempt = 1; attempt <= retries; attempt++) {
        try {
            await workbook.xlsx.writeFile(EXCEL_FILE);
            return;
        } catch (error) {
            if (
                !["EBUSY", "EPERM", "EACCES"].includes(error.code) ||
                attempt === retries
            ) {
                throw error;
            }

            console.log(`Excel locked. Retry ${attempt}/${retries}`);
            await new Promise(resolve => setTimeout(resolve, 1000 * attempt));
        }
    }
}

// Create or load Excel workbook
async function getWorkbook() {
    const workbook = new ExcelJS.Workbook();

    if (fs.existsSync(EXCEL_FILE) && fs.statSync(EXCEL_FILE).size > 0) {
        await workbook.xlsx.readFile(EXCEL_FILE);
    }

    let ordersSheet = workbook.getWorksheet("Orders");
    let summarySheet = workbook.getWorksheet("Summary");

    if (!ordersSheet) {
        ordersSheet = workbook.addWorksheet("Orders");
    }

    if (!summarySheet) {
        summarySheet = workbook.addWorksheet("Summary");
    }

    // Headers for Orders sheet
    if (ordersSheet.rowCount === 0) {
        ordersSheet.addRow([
            "Order ID",
            "Date & Time",
            "Product Name",
            "Category",
            "Pack Size",
            "Quantity",
            "Price/Unit (₹)",
            "Item Total (₹)",
            "Order Total (₹)"
        ]);

        ordersSheet.columns = [
            { width: 24 },
            { width: 23 },
            { width: 28 },
            { width: 18 },
            { width: 14 },
            { width: 12 },
            { width: 17 },
            { width: 17 },
            { width: 18 }
        ];

        styleHeader(ordersSheet);
    }

    // Headers for Summary sheet
    if (summarySheet.rowCount === 0) {
        summarySheet.addRow([
            "Order ID",
            "Date & Time",
            "Items Count",
            "Grand Total (₹)"
        ]);

        summarySheet.columns = [
            { width: 25 },
            { width: 23 },
            { width: 16 },
            { width: 20 }
        ];

        styleHeader(summarySheet);
    }

    return { workbook, ordersSheet, summarySheet };
}

// Style header
function styleHeader(sheet) {
    const header = sheet.getRow(1);

    header.font = {
        bold: true,
        color: { argb: "FFFFFFFF" },
        size: 11
    };

    header.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF2D6A4F" }
    };

    header.alignment = {
        horizontal: "center",
        vertical: "middle"
    };

    header.height = 25;

    header.eachCell(cell => {
        cell.border = {
            top: { style: "thin" },
            bottom: { style: "thin" },
            left: { style: "thin" },
            right: { style: "thin" }
        };
    });
}

// ==============================
// POST /api/order
// Save new order
// ==============================

app.post("/api/order", async (req, res) => {
    try {
        const order = req.body;

        if (
            !order ||
            !Array.isArray(order.items) ||
            order.items.length === 0
        ) {
            return res.status(400).json({
                success: false,
                error: "Invalid order data"
            });
        }

        const orderId = order.orderId || `ORD-${Date.now()}`;
        const orderDate = order.date || new Date().toLocaleString("en-IN");

        const items = order.items.map(item => {
            const quantity = Number(item.quantity);
            const price = Number(item.pricePerUnit);

            if (
                !item.name ||
                !Number.isFinite(quantity) ||
                !Number.isFinite(price) ||
                quantity <= 0 ||
                price < 0
            ) {
                throw new Error("Invalid product or quantity/price");
            }

            return {
                name: item.name,
                category: item.category || "",
                unit: item.unit || "",
                quantity,
                price,
                total: Number.isFinite(Number(item.total))
                    ? Number(item.total)
                    : quantity * price
            };
        });

        const grandTotal = Number.isFinite(Number(order.grandTotal))
            ? Number(order.grandTotal)
            : items.reduce((sum, item) => sum + item.total, 0);

        await queueSave(async () => {
            const {
                workbook,
                ordersSheet,
                summarySheet
            } = await getWorkbook();

            // Prevent duplicate order IDs
            let duplicate = false;

            summarySheet.eachRow((row, rowNumber) => {
                if (
                    rowNumber > 1 &&
                    String(row.getCell(1).value) === String(orderId)
                ) {
                    duplicate = true;
                }
            });

            if (duplicate) {
                throw new Error(`Duplicate order ID: ${orderId}`);
            }

            // Add every product to Orders sheet
            for (const item of items) {
                const row = ordersSheet.addRow([
                    orderId,
                    orderDate,
                    item.name,
                    item.category,
                    item.unit,
                    item.quantity,
                    item.price,
                    item.total,
                    grandTotal
                ]);

                row.eachCell(cell => {
                    cell.alignment = {
                        horizontal: "center",
                        vertical: "middle"
                    };

                    cell.border = {
                        top: { style: "thin" },
                        bottom: { style: "thin" },
                        left: { style: "thin" },
                        right: { style: "thin" }
                    };
                });

                if (row.number % 2 === 0) {
                    row.eachCell(cell => {
                        cell.fill = {
                            type: "pattern",
                            pattern: "solid",
                            fgColor: { argb: "FFD8F3DC" }
                        };
                    });
                }
            }

            // Add one entry per order to Summary sheet
            summarySheet.addRow([
                orderId,
                orderDate,
                items.reduce((sum, item) => sum + item.quantity, 0),
                grandTotal
            ]);

            // Save Excel file
            await saveWorkbook(workbook);

            console.log("✅ Order saved:", orderId);
            console.log("📁 Excel file:", EXCEL_FILE);
        });

        res.json({
            success: true,
            message: "Order saved successfully!",
            orderId,
            file: EXCEL_FILE
        });

    } catch (error) {
        console.error("❌ Error saving order:", error);

        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// ==============================
// GET /api/orders
// Get all orders
// ==============================

app.get("/api/orders", async (req, res) => {
    try {
        await saveQueue;

        if (
            !fs.existsSync(EXCEL_FILE) ||
            fs.statSync(EXCEL_FILE).size === 0
        ) {
            return res.json({ success: true, orders: [] });
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(EXCEL_FILE);

        const sheet = workbook.getWorksheet("Summary");

        if (!sheet) {
            return res.json({ success: true, orders: [] });
        }

        const orders = [];

        sheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return;

            orders.push({
                orderId: row.getCell(1).value,
                date: row.getCell(2).value,
                itemCount: row.getCell(3).value,
                grandTotal: row.getCell(4).value
            });
        });

        res.json({
            success: true,
            totalOrders: orders.length,
            orders
        });

    } catch (error) {
        console.error("❌ Error reading orders:", error);

        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// ==============================
// GET /api/products
// Product list
// ==============================

app.get("/api/products", (req, res) => {
    const products = [
        { id: 1, name: "Basmati Rice", category: "grains", price: 85, unit: "1 kg", stock: 50 },
        { id: 2, name: "Wheat Flour (Atta)", category: "grains", price: 45, unit: "1 kg", stock: 40 },
        { id: 3, name: "Toor Dal", category: "grains", price: 130, unit: "500 g", stock: 30 },
        { id: 4, name: "Chana Dal", category: "grains", price: 95, unit: "500 g", stock: 25 },
        { id: 5, name: "Poha", category: "grains", price: 55, unit: "500 g", stock: 20 },
        { id: 6, name: "Sunflower Oil", category: "oil", price: 145, unit: "1 L", stock: 20 },
        { id: 7, name: "Mustard Oil", category: "oil", price: 155, unit: "1 L", stock: 18 },
        { id: 8, name: "Pure Desi Ghee", category: "oil", price: 480, unit: "500 ml", stock: 10 },
        { id: 9, name: "Coconut Oil", category: "oil", price: 180, unit: "500 ml", stock: 15 },
        { id: 10, name: "Red Chilli Powder", category: "spices", price: 65, unit: "200 g", stock: 35 },
        { id: 11, name: "Turmeric Powder", category: "spices", price: 55, unit: "200 g", stock: 40 },
        { id: 12, name: "Cumin Seeds", category: "spices", price: 85, unit: "200 g", stock: 30 },
        { id: 13, name: "Garam Masala", category: "spices", price: 75, unit: "100 g", stock: 25 },
        { id: 14, name: "Coriander Powder", category: "spices", price: 60, unit: "200 g", stock: 30 },
        { id: 15, name: "Full Cream Milk", category: "dairy", price: 60, unit: "1 L", stock: 20 },
        { id: 16, name: "Paneer", category: "dairy", price: 90, unit: "200 g", stock: 15 },
        { id: 17, name: "Curd (Dahi)", category: "dairy", price: 45, unit: "400 g", stock: 25 },
        { id: 18, name: "Butter", category: "dairy", price: 55, unit: "100 g", stock: 20 },
        { id: 19, name: "Namkeen Mixture", category: "snacks", price: 40, unit: "200 g", stock: 30 },
        { id: 20, name: "Biscuits (Marie)", category: "snacks", price: 25, unit: "200 g", stock: 50 },
        { id: 21, name: "Papad", category: "snacks", price: 35, unit: "200 g", stock: 25 },
        { id: 22, name: "Roasted Peanuts", category: "snacks", price: 30, unit: "200 g", stock: 40 },
        { id: 23, name: "Tea Leaves (Chai)", category: "beverages", price: 90, unit: "250 g", stock: 30 },
        { id: 24, name: "Coffee Powder", category: "beverages", price: 120, unit: "100 g", stock: 20 },
        { id: 25, name: "Sugar", category: "beverages", price: 45, unit: "1 kg", stock: 40 },
        { id: 26, name: "Jaggery (Gur)", category: "beverages", price: 60, unit: "500 g", stock: 25 },
        { id: 27, name: "Surf Excel Detergent", category: "cleaning", price: 85, unit: "500 g", stock: 20 },
        { id: 28, name: "Vim Bar", category: "cleaning", price: 20, unit: "1 bar", stock: 40 },
        { id: 29, name: "Phenyl", category: "cleaning", price: 55, unit: "500 ml", stock: 15 },
        { id: 30, name: "Toilet Cleaner", category: "cleaning", price: 65, unit: "500 ml", stock: 12 }
    ];

    res.json({ success: true, products });
});

// ==============================
// GET /api/health
// Check backend and Excel file
// ==============================

app.get("/api/health", (req, res) => {
    res.json({
        success: true,
        server: "Running",
        excelFile: EXCEL_FILE,
        excelExists: fs.existsSync(EXCEL_FILE)
    });
});

// ==============================
// Start Server
// ==============================

app.listen(PORT, () => {
    console.log("\n🛒 KiranaShop Backend Started!");
    console.log(`🌐 Website: http://localhost:${PORT}`);
    console.log(`📊 Orders: http://localhost:${PORT}/api/orders`);
    console.log(`📁 Excel file: ${EXCEL_FILE}`);
    console.log(`❤️ Health: http://localhost:${PORT}/api/health\n`);
});
