
const express = require("express");
const cors = require("cors");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const app = express();

// Render automatically provides PORT
const PORT = process.env.PORT || 3000;

// Excel storage location
// Render Persistent Disk ke liye DATA_DIR=/var/data set karna
const DATA_DIR = process.env.DATA_DIR || __dirname;
const EXCEL_FILE = path.join(DATA_DIR, "orders.xlsx");

if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
}

// ---------- CORS CONFIGURATION ----------
const allowedOrigins = (process.env.FRONTEND_URL || "")
    .split(",")
    .map(origin => origin.trim())
    .filter(Boolean);

app.use(cors({
    origin: (origin, callback) => {
        // Allow requests without browser Origin (health checks, etc.)
        if (!origin) return callback(null, true);

        // Allow all origins until FRONTEND_URL is configured
        if (allowedOrigins.length === 0) {
            return callback(null, true);
        }

        if (allowedOrigins.includes(origin)) {
            return callback(null, true);
        }

        return callback(new Error("Not allowed by CORS"));
    }
}));

app.use(express.json({ limit: "1mb" }));

// ---------- EXCEL SAVE QUEUE ----------
let saveQueue = Promise.resolve();

function queueSave(task) {
    const current = saveQueue.then(task, task);
    saveQueue = current.catch(() => {});
    return current;
}

// ---------- SAVE WORKBOOK ----------
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
            await new Promise(resolve =>
                setTimeout(resolve, 1000 * attempt)
            );
        }
    }
}

// ---------- HEADER STYLE ----------
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

// ---------- GET WORKBOOK ----------
async function getWorkbook() {
    const workbook = new ExcelJS.Workbook();

    if (
        fs.existsSync(EXCEL_FILE) &&
        fs.statSync(EXCEL_FILE).size > 0
    ) {
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

// ---------- HOME ROUTE ----------
app.get("/", (req, res) => {
    res.send("KiranaShop Backend is Running! 🛒");
});

// ---------- HEALTH CHECK ----------
app.get("/api/health", (req, res) => {
    res.json({
        success: true,
        server: "Running",
        platform: "Render",
        excelFile: EXCEL_FILE,
        excelExists: fs.existsSync(EXCEL_FILE),
        storage: process.env.DATA_DIR ? "Configured DATA_DIR" : "Temporary local storage"
    });
});

// ---------- SAVE ORDER ----------
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

        if (order.items.length > 100) {
            return res.status(400).json({
                success: false,
                error: "Too many order items"
            });
        }

        const orderId = order.orderId || `ORD-${Date.now()}`;
        const orderDate = new Date().toLocaleString("en-IN");

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
                throw new Error("Invalid product, quantity or price");
            }

            return {
                name: String(item.name),
                category: String(item.category || ""),
                unit: String(item.unit || ""),
                quantity,
                price,
                total: quantity * price
            };
        });

        // Calculate total on backend
        const grandTotal = items.reduce(
            (sum, item) => sum + item.total,
            0
        );

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
                throw new Error("Duplicate order ID");
            }

            // Save product details
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

            // Save order summary
            summarySheet.addRow([
                orderId,
                orderDate,
                items.reduce((sum, item) => sum + item.quantity, 0),
                grandTotal
            ]);

            await saveWorkbook(workbook);

            console.log("Order saved:", orderId);
            console.log("Excel file:", EXCEL_FILE);
        });

        res.status(201).json({
            success: true,
            message: "Order saved successfully!",
            orderId,
            grandTotal
        });

    } catch (error) {
        console.error("Order Error:", error);

        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// ---------- GET ALL ORDERS ----------
app.get("/api/orders", async (req, res) => {
    try {
        await saveQueue;

        if (
            !fs.existsSync(EXCEL_FILE) ||
            fs.statSync(EXCEL_FILE).size === 0
        ) {
            return res.json({
                success: true,
                totalOrders: 0,
                orders: []
            });
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(EXCEL_FILE);

        const sheet = workbook.getWorksheet("Summary");

        if (!sheet) {
            return res.json({
                success: true,
                totalOrders: 0,
                orders: []
            });
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
        console.error("Read Orders Error:", error);

        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// ---------- PRODUCTS API ----------
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

    res.json({
        success: true,
        products
    });
});

// ---------- 404 HANDLER ----------
app.use((req, res) => {
    res.status(404).json({
        success: false,
        error: "API route not found"
    });
});

// ---------- ERROR HANDLER ----------
app.use((err, req, res, next) => {
    console.error("Server Error:", err.message);

    res.status(500).json({
        success: false,
        error: err.message
    });
});

// ---------- START SERVER ----------
app.listen(PORT, "0.0.0.0", () => {
    console.log("KiranaShop Backend Started!");
    console.log(`Port: ${PORT}`);
    console.log(`Excel: ${EXCEL_FILE}`);
    console.log(`Health: http://localhost:${PORT}/api/health`);
});
