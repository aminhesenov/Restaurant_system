const defaultProducts = [
    { id: 1, name: "Yarpaq Dolması", price: 8.50, category: "qazan" },
    { id: 2, name: "Dana Basdırma", price: 14.00, category: "manqal" },
    { id: 3, name: "Lülə Kabab", price: 10.00, category: "manqal" },
    { id: 4, name: "Mərci Şorbası", price: 4.50, category: "sorba" },
    { id: 5, name: "Dovğa", price: 4.00, category: "sorba" },
    { id: 6, name: "Sezar Salatı", price: 9.00, category: "salat" },
    { id: 7, name: "Çoban Salatı", price: 5.00, category: "salat" },
    { id: 8, name: "Meyvə Assorti", price: 12.00, category: "meyve" },
    { id: 9, name: "Pomidor-Xiyar", price: 4.00, category: "meyve" },
    { id: 10, name: "Coca-Cola 330ml", price: 2.50, category: "ickiler" },
    { id: 11, name: "Ev Sayı Limonad", price: 5.00, category: "ickiler" }
];

const defaultTables = [
    { id: 1, number: "Masa 1", status: "empty", orders: [], amountPaid: 0 },
    { id: 2, number: "Masa 2", status: "empty", orders: [], amountPaid: 0 },
    { id: 3, number: "Masa 3", status: "empty", orders: [], amountPaid: 0 },
    { id: 4, number: "Masa 4", status: "empty", orders: [], amountPaid: 0 },
    { id: 5, number: "Masa 5", status: "empty", orders: [], amountPaid: 0 },
    { id: 6, number: "Masa 6", status: "empty", orders: [], amountPaid: 0 },
    { id: 7, number: "Masa 7", status: "empty", orders: [], amountPaid: 0 },
    { id: 8, number: "Masa 8", status: "empty", orders: [], amountPaid: 0 },
    { id: 9, number: "Masa 9", status: "empty", orders: [], amountPaid: 0 }
];

let products = JSON.parse(localStorage.getItem('rest_products'));
if (!products || products.length === 0) {
    products = defaultProducts;
    localStorage.setItem('rest_products', JSON.stringify(products));
}

let tables = JSON.parse(localStorage.getItem('rest_tables'));
if (!tables || tables.length === 0) {
    tables = defaultTables;
    localStorage.setItem('rest_tables', JSON.stringify(tables));
}

let salesArchive = JSON.parse(localStorage.getItem('rest_sales_archive')) || [];

function updateStorage() {
    localStorage.setItem('rest_products', JSON.stringify(products));
    localStorage.setItem('rest_tables', JSON.stringify(tables));
    localStorage.setItem('rest_sales_archive', JSON.stringify(salesArchive));
}

function calculateTotalEarnings() {
    let total = salesArchive.reduce((sum, sale) => {
        let amount = sale.totalAmount ? parseFloat(sale.totalAmount) : 0;
        return isNaN(amount) ? sum : sum + amount;
    }, 0);

    const earningsElement = document.getElementById('totalEarningsVal') || document.getElementById('totalEarnings');
    if (earningsElement) {
        earningsElement.innerText = `${total.toFixed(2)} AZN`;
    }
}

function exportToExcel() {
    try {
        if (!salesArchive || salesArchive.length === 0) {
            salesArchive = JSON.parse(localStorage.getItem('rest_sales_archive')) || [];
        }

        if (salesArchive.length === 0) {
            alert("Diqqət: Arxiviniz hazırda boşdur! Hesabat yaradıla bilmədi.");
            return;
        }

        let excelTemplate = `
                    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
                    <head>
                        <meta charset="utf-8">
                        <style>
                            body, table { background-color: #ffffff; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
                            td { background-color: #ffffff; border: none; padding: 6px; font-size: 11pt; vertical-align: middle; }
                            .th-header { background-color: #2c3e50; color: #ffffff; font-weight: bold; text-align: left; height: 30px; }
                            .text-cell { mso-number-format:"\\@"; text-align: left; }
                            .number-cell { mso-number-format:"\\#,##0\\.00"; text-align: right; }
                            .total-row-label { font-weight: bold; background-color: #ffffff; font-size: 12pt; text-align: left; }
                            .total-row-value { font-weight: bold; background-color: #ffffff; font-size: 12pt; text-align: right; border-top: 2px solid #2c3e50; color: #27ae60; }
                        </style>
                    </head>
                    <body>
                        <table>
                            <thead>
                                <tr>
                                    <th class="th-header" style="width: 100px;">Sifariş ID</th>
                                    <th class="th-header" style="width: 120px;">Masa</th>
                                    <th class="th-header" style="width: 160px;">Tarix</th>
                                    <th class="th-header" style="width: 150px;">Ümumi Məbləğ</th>
                                    <th class="th-header" style="width: 130px;">Ödənilən</th>
                                    <th class="th-header" style="width: 110px;">Qalıq</th>
                                    <th class="th-header" style="width: 350px;">Sifariş Tərkibi</th>
                                </tr>
                            </thead>
                            <tbody>
                `;

        let totalEarnings = 0;
        let validRowCount = 0;

        salesArchive.forEach(sale => {
            if (!sale.totalAmount || isNaN(sale.totalAmount)) {
                return;
            }

            validRowCount++;
            let amt = parseFloat(sale.totalAmount);
            totalEarnings += amt;

            let saleId = sale.id ? sale.id : "N/A";
            let tableNum = sale.tableNumber ? sale.tableNumber : (sale.tableName ? sale.tableName : "Bilinməyən Masa");
            let saleDate = sale.date ? sale.date : "Tarix qeyd edilməyib";
            let amountPaid = sale.amountPaid ? parseFloat(sale.amountPaid) : 0;
            let changeAmount = sale.changeAmount ? parseFloat(sale.changeAmount) : 0;

            let itemsSummary = "Məhsul yoxdur";
            if (sale.items && Array.isArray(sale.items) && sale.items.length > 0) {
                itemsSummary = sale.items.map(i => `${i.name || 'Naməlum'} (${i.quantity || 1} ədəd)`).join(" | ");
            }

            excelTemplate += `
                        <tr>
                            <td class="text-cell">${saleId}</td>
                            <td class="text-cell">${tableNum}</td>
                            <td class="text-cell">${saleDate}</td>
                            <td class="number-cell">${amt}</td>
                            <td class="number-cell">${amountPaid}</td>
                            <td class="number-cell">${changeAmount}</td>
                            <td class="text-cell">${itemsSummary}</td>
                        </tr>
                    `;
        });

        if (validRowCount === 0) {
            alert("Arxivdə etibarlı satış məlumatı tapılmadı!");
            return;
        }

        excelTemplate += `
                            <tr><td colspan="7" style="height: 15px; background-color: #ffffff;"></td></tr>
                            <tr>
                                <td class="total-row-label" colspan="3">YEKUN QAZANC</td>
                                <td class="total-row-value">${totalEarnings}</td>
                                <td colspan="3" style="background-color: #ffffff;"></td>
                            </tr>
                        </tbody>
                    </table>
                    </body>
                    </html>
                `;

        const blob = new Blob([excelTemplate], { type: "application/vnd.ms-excel;charset=utf-8;" });
        const url = URL.createObjectURL(blob);

        const link = document.createElement("a");
        link.href = url;
        link.setAttribute("download", `Restoran_Satis_Hesabati_${new Date().toISOString().slice(0, 10)}.xls`);

        document.body.appendChild(link);
        link.click();

        document.body.removeChild(link);
        URL.revokeObjectURL(url);

    } catch (error) {
        alert("Excel yüklənərkən daxili xəta baş verdi: " + error.message);
    }
}

let currentTableId = null;
let isEditModeActive = false;
let selectedCategory = "all";
let searchQuery = "";

function switchCategory(category) {
    selectedCategory = category;
    const tabs = document.querySelectorAll('.category-tab');
    tabs.forEach(tab => tab.classList.remove('active-tab'));
    if (event && event.target) event.target.classList.add('active-tab');
    renderMenu();
}

function handleSearchInput() {
    const inputElement = document.getElementById('menuSearchInput');
    if (inputElement) {
        searchQuery = inputElement.value.trim().toLowerCase();
        renderMenu();
    }
}

function toggleMenuForm() {
    const formPanel = document.getElementById('addProductFormPanel');
    const toggleBtn = document.getElementById('toggleFormBtn');
    if (!formPanel || !toggleBtn) return;

    const currentDisplay = window.getComputedStyle(formPanel).display;

    const editIdForm = document.getElementById('editProductId');
    if (currentDisplay === 'block' && (!editIdForm || !editIdForm.value)) {
        formPanel.style.display = 'none';
        toggleBtn.innerText = "Menyunu Redaktə Et";
        toggleBtn.classList.remove('active');
        isEditModeActive = false;
    } else {
        resetForm();
        formPanel.style.display = 'block';
        toggleBtn.innerText = "Paneli Bağla";
        toggleBtn.classList.add('active');
        isEditModeActive = true;
    }
    renderMenu();
}

function resetForm() {
    const eId = document.getElementById('editProductId');
    const pName = document.getElementById('newProductName');
    const pPrice = document.getElementById('newProductPrice');
    const pCat = document.getElementById('newProductCategory');
    const pTitle = document.getElementById('formPanelTitle');
    const pBtn = document.getElementById('formSubmitBtn');

    if (eId) eId.value = "";
    if (pName) pName.value = "";
    if (pPrice) pPrice.value = "";
    if (pCat) pCat.value = "qazan";
    if (pTitle) pTitle.innerText = "Menyuya Yeni Yemək Əlavə Et";
    if (pBtn) pBtn.innerText = "Menyuya Əlavə Et";
}

function saveProduct() {
    const nameInput = document.getElementById('newProductName');
    const priceInput = document.getElementById('newProductPrice');
    const catInput = document.getElementById('newProductCategory');
    const editIdInput = document.getElementById('editProductId');

    if (!nameInput || !priceInput || !catInput) return;

    const name = nameInput.value.trim();
    const price = parseFloat(priceInput.value);
    const category = catInput.value;

    if (name === "" || isNaN(price) || price <= 0) {
        alert("Zəhmət olmasa düzgün yemək adı və qiymət daxil edin!");
        return;
    }

    if (editIdInput && editIdInput.value) {
        const prodId = parseInt(editIdInput.value);
        const product = products.find(p => p.id === prodId);
        if (product) {
            product.name = name;
            product.price = price;
            product.category = category;
            tables.forEach(t => {
                const order = t.orders.find(o => o.productId === prodId);
                if (order) { order.name = name; order.price = price; }
            });
        }
        alert("Məhsul uğurla yeniləndi!");
    } else {
        const maxId = products.length > 0 ? Math.max(...products.map(p => p.id)) : 0;
        products.push({ id: maxId + 1, name: name, price: price, category: category });
        alert(`"${name}" uğurla menyuya əlavə edildi!`);
    }

    updateStorage();
    resetForm();
    renderMenu();
    renderOrderDetails();
}

function editProduct(productId) {
    const product = products.find(p => p.id === productId);
    if (!product) return;

    const formPanel = document.getElementById('addProductFormPanel');
    if (formPanel) formPanel.style.display = 'block';

    const eId = document.getElementById('editProductId');
    const pName = document.getElementById('newProductName');
    const pPrice = document.getElementById('newProductPrice');
    const pCat = document.getElementById('newProductCategory');
    const pTitle = document.getElementById('formPanelTitle');
    const pBtn = document.getElementById('formSubmitBtn');

    if (eId) eId.value = product.id;
    if (pName) pName.value = product.name;
    if (pPrice) pPrice.value = product.price;
    if (pCat) pCat.value = product.category || "qazan";
    if (pTitle) pTitle.innerText = "Məhsulu Redaktə Et";
    if (pBtn) pBtn.innerText = "Dəyişikliyi Yadda Saxla";
}

function deleteProduct(productId) {
    const product = products.find(p => p.id === productId);
    if (!product) return;

    const hasActiveOrder = tables.some(t => t.orders.some(o => o.productId === productId));
    if (hasActiveOrder) {
        alert("Bu məhsul hazırda masalardan birində sifariş edildiyi üçün silinə bilməz!");
        return;
    }

    if (confirm(`"${product.name}" menyudan silinsin?`)) {
        const index = products.findIndex(p => p.id === productId);
        if (index > -1) {
            products.splice(index, 1);
            const editIdInput = document.getElementById('editProductId');
            if (editIdInput && editIdInput.value == productId) resetForm();
            updateStorage();
            renderMenu();
        }
    }
}

function addTable() {
    const nextId = tables.length > 0 ? Math.max(...tables.map(t => t.id)) + 1 : 1;
    tables.push({ id: nextId, number: `Masa ${nextId}`, status: "empty", orders: [], amountPaid: 0 });
    updateStorage();
    renderTables();
}

function removeTable() {
    if (tables.length === 0) return;
    const lastTable = tables[tables.length - 1];
    if (lastTable.status === 'occupied') {
        alert("Aktiv sifarişi olan masa silinə bilməz!");
        return;
    }
    if (confirm(`"${lastTable.number}" silinsin?`)) {
        if (currentTableId === lastTable.id) currentTableId = null;
        tables.pop();
        updateStorage();
        renderTables();
        renderOrderDetails();
    }
}

function renderTables() {
    const grid = document.getElementById('tablesGrid');
    if (!grid) return;
    grid.innerHTML = '';

    const kitchenOrders = JSON.parse(localStorage.getItem('rest_kitchen_orders')) || [];

    tables.forEach(table => {
        const card = document.createElement('div');
        card.className = `table-card ${table.status === 'occupied' ? 'occupied' : ''} ${table.id === currentTableId ? 'active-table' : ''}`;
        card.onclick = () => selectTable(table.id);

        const currentKitchenOrder = kitchenOrders.find(ko => ko.tableName === table.number);

        let statusBadge = '';
        if (table.status === 'occupied') {
            if (currentKitchenOrder) {
                if (currentKitchenOrder.status === 'ready') {
                    statusBadge = `<div style="font-size: 11px; font-weight: bold; color: #27ae60; margin-top: 4px;">Hazırdır ✔️</div>`;
                } else if (currentKitchenOrder.status === 'pending') {
                    statusBadge = `<div style="font-size: 11px; font-weight: bold; color: #e67e22; margin-top: 4px;">Hazırlanır ⏳</div>`;
                }
            } else {
                statusBadge = `<div style="font-size: 11px; color: #95a5a6; margin-top: 4px;">Gözləyir ⏳</div>`;
            }
        }

        card.innerHTML = `
            <div class="table-number">${table.number}</div>
            <div class="table-status">${table.status === 'occupied' ? 'Dolu' : 'Boş'}</div>
            ${statusBadge}
        `;
        grid.appendChild(card);
    });
}

function renderMenu() {
    const menuGrid = document.getElementById('menuGrid');
    if (!menuGrid) return;
    menuGrid.innerHTML = '';

    const filteredProducts = products.filter(p => {
        const matchesCategory = (selectedCategory === "all" || p.category === selectedCategory);
        const matchesSearch = p.name.toLowerCase().includes(searchQuery);
        return matchesCategory && matchesSearch;
    });

    if (filteredProducts.length === 0) {
        menuGrid.innerHTML = `<p style="color: #95a5a6; padding: 10px; text-align:center;">Uyğun məhsul tapılmadı.</p>`;
        return;
    }

    filteredProducts.forEach(product => {
        const row = document.createElement('div');
        row.className = 'menu-item-row';

        const clickPart = document.createElement('div');
        clickPart.className = 'menu-item-clickable';
        clickPart.innerText = `${product.name} - ${product.price.toFixed(2)} AZN`;
        clickPart.onclick = () => addProductToOrder(product.id);

        const managementDiv = document.createElement('div');
        managementDiv.className = 'menu-management-btns';
        managementDiv.style.display = isEditModeActive ? 'flex' : 'none';

        const editBtn = document.createElement('button');
        editBtn.className = 'action-btn';
        editBtn.innerText = 'Redaktə';
        editBtn.style.marginRight = '5px';
        editBtn.onclick = (e) => { e.stopPropagation(); editProduct(product.id); };

        const deleteBtn = document.createElement('button');
        deleteBtn.className = 'action-btn';
        deleteBtn.innerText = 'Sil';
        deleteBtn.onclick = (e) => { e.stopPropagation(); deleteProduct(product.id); };

        managementDiv.appendChild(editBtn);
        managementDiv.appendChild(deleteBtn);

        row.appendChild(clickPart);
        row.appendChild(managementDiv);
        menuGrid.appendChild(row);
    });
}

function selectTable(tableId) {
    currentTableId = tableId;
    const table = tables.find(t => t.id === tableId);

    const titleEl = document.getElementById('selectedTableTitle');
    if (titleEl) titleEl.innerHTML = `${table.number} - Sifariş Paneli`;

    const pArea = document.getElementById('paymentArea');
    const btnCheck = document.getElementById('checkoutBtn');
    const btnRec = document.getElementById('receiptBtn');
    const btnKit = document.getElementById('kitchenBtn');
    const inputPaid = document.getElementById('amountPaidInput');

    if (table.orders.length > 0) {
        if (pArea) pArea.style.display = 'block';
        if (btnCheck) btnCheck.disabled = false;
        if (btnRec) btnRec.disabled = false;
        if (btnKit) btnKit.disabled = false;
        if (inputPaid) inputPaid.value = table.amountPaid || table.orders.reduce((sum, item) => sum + (item.price * item.quantity), 0);
        calculateChange();
    } else {
        if (pArea) pArea.style.display = 'none';
        if (btnCheck) btnCheck.disabled = true;
        if (btnRec) btnRec.disabled = true;
        if (btnKit) btnKit.disabled = true;
    }
    renderTables();
    renderOrderDetails();
}

function addProductToOrder(productId) {
    if (!currentTableId) {
        alert("Zəhmət olmasa, əvvəlcə masa seçin!");
        return;
    }
    const table = tables.find(t => t.id === currentTableId);
    const product = products.find(p => p.id === productId);
    const existingOrderIndex = table.orders.findIndex(o => o.productId === productId);

    if (existingOrderIndex > -1) {
        table.orders[existingOrderIndex].quantity += 1;
    } else {
        table.orders.push({ productId: product.id, name: product.name, price: product.price, quantity: 1 });
    }

    table.status = 'occupied';

    const pArea = document.getElementById('paymentArea');
    const btnCheck = document.getElementById('checkoutBtn');
    const btnRec = document.getElementById('receiptBtn');
    const btnKit = document.getElementById('kitchenBtn');
    const inputPaid = document.getElementById('amountPaidInput');

    if (pArea) pArea.style.display = 'block';
    if (btnCheck) btnCheck.disabled = false;
    if (btnRec) btnRec.disabled = false;
    if (btnKit) btnKit.disabled = false;

    let tableTotal = table.orders.reduce((sum, item) => sum + (item.price * item.quantity), 0);
    if (inputPaid) inputPaid.value = tableTotal;

    updateStorage();
    renderTables();
    renderOrderDetails();
}

function changeQuantity(productId, amount) {
    const table = tables.find(t => t.id === currentTableId);
    const orderItem = table.orders.find(o => o.productId === productId);

    if (orderItem) {
        orderItem.quantity += amount;
        if (orderItem.quantity <= 0) {
            table.orders = table.orders.filter(o => o.productId !== productId);
        }
    }

    const pArea = document.getElementById('paymentArea');
    const btnCheck = document.getElementById('checkoutBtn');
    const btnRec = document.getElementById('receiptBtn');
    const btnKit = document.getElementById('kitchenBtn');
    const inputPaid = document.getElementById('amountPaidInput');

    if (table.orders.length === 0) {
        table.status = 'empty';
        if (pArea) pArea.style.display = 'none';
        if (btnCheck) btnCheck.disabled = true;
        if (btnRec) btnRec.disabled = true;
        if (btnKit) btnKit.disabled = true;
    } else {
        let tableTotal = table.orders.reduce((sum, item) => sum + (item.price * item.quantity), 0);
        if (inputPaid) inputPaid.value = tableTotal;
    }
    updateStorage();
    renderTables();
    renderOrderDetails();
}

function calculateChange() {
    if (!currentTableId) return;
    const table = tables.find(t => t.id === currentTableId);
    let tableTotal = table.orders.reduce((sum, item) => sum + (item.price * item.quantity), 0);

    const inputPaid = document.getElementById('amountPaidInput');
    const amountPaid = inputPaid ? (parseFloat(inputPaid.value) || 0) : 0;
    table.amountPaid = amountPaid;
    updateStorage();

    const change = amountPaid - tableTotal;
    const changeContainer = document.getElementById('changeAmount');

    if (changeContainer) {
        if (change >= 0) {
            changeContainer.innerText = `${change.toFixed(2)} AZN`;
            changeContainer.style.color = '#2ecc71';
        } else {
            changeContainer.innerText = `Çatışmır: ${Math.abs(change).toFixed(2)} AZN`;
            changeContainer.style.color = '#e74c3c';
        }
    }
}

function renderOrderDetails() {
    const listContainer = document.getElementById('orderItemsList');
    const totalAmtEl = document.getElementById('totalAmount');
    const titleEl = document.getElementById('selectedTableTitle');

    if (!listContainer) return;

    if (!currentTableId) {
        listContainer.innerHTML = `<p style="color: #95a5a6; text-align: center; margin-top: 50px;">Sifarişə başlamaq üçün masa seçin.</p>`;
        if (totalAmtEl) totalAmtEl.innerText = "0.00 AZN";
        if (titleEl) titleEl.innerText = "Masa seçilməyib";
        return;
    }

    const table = tables.find(t => t.id === currentTableId);

    const kitchenOrders = JSON.parse(localStorage.getItem('rest_kitchen_orders')) || [];
    const currentKitchenOrder = kitchenOrders.find(ko => ko.tableName === table.number);

    let statusBadge = '<span style="color: #95a5a6; font-size: 13px; font-weight: normal; margin-left: 10px;">Gözləyir ⏳</span>';
    if (currentKitchenOrder) {
        if (currentKitchenOrder.status === 'ready') {
            statusBadge = '<span style="color: #27ae60; font-weight: bold; font-size: 13px; margin-left: 10px;">Hazırdır ✔️</span>';
        } else if (currentKitchenOrder.status === 'pending') {
            statusBadge = '<span style="color: #e67e22; font-weight: bold; font-size: 13px; margin-left: 10px;">Hazırlanır ⏳</span>';
        }
    }
    if (titleEl) {
        titleEl.innerHTML = `${table.number} - Sifariş Paneli ${statusBadge}`;
    }

    if (table.orders.length === 0) {
        listContainer.innerHTML = `<p style="color: #95a5a6; text-align: center; margin-top: 50px;">Sifariş boşdur. Menyudan yemək seçin.</p>`;
        if (totalAmtEl) totalAmtEl.innerText = "0.00 AZN";
        return;
    }

    listContainer.innerHTML = '';
    let grandTotal = 0;

    table.orders.forEach(item => {
        const totalItemPrice = item.price * item.quantity;
        grandTotal += totalItemPrice;

        const row = document.createElement('div');
        row.className = 'order-item-row';
        row.innerHTML = `
                    <div class="item-details">
                        <div class="item-name">${item.name}</div>
                        <div class="item-price">${item.price.toFixed(2)} x ${item.quantity} = ${totalItemPrice.toFixed(2)} AZN</div>
                    </div>
                    <div class="quantity-controls">
                        <button class="action-btn" style="background-color:#e74c3c; padding:2px 8px;" onclick="changeQuantity(${item.productId}, -1)">-</button>
                        <span style="font-weight:bold; min-width:20px; text-align:center;">${item.quantity}</span>
                        <button class="action-btn" style="background-color:#2ecc71; padding:2px 8px;" onclick="changeQuantity(${item.productId}, 1)">+</button>
                    </div>
                `;
        listContainer.appendChild(row);
    });

    if (totalAmtEl) totalAmtEl.innerText = `${grandTotal.toFixed(2)} AZN`;
}

function sendToKitchen() {
    if (!currentTableId) return;
    const table = tables.find(t => t.id === currentTableId);
    if (table.orders.length === 0) return;

    let kitchenOrders = JSON.parse(localStorage.getItem('rest_kitchen_orders')) || [];
    kitchenOrders = kitchenOrders.filter(ko => !(ko.tableName === table.number && ko.status === 'pending'));

    const newKitchenOrder = {
        id: Math.floor(1000 + Math.random() * 9000),
        tableName: table.number,
        time: new Date().toLocaleTimeString('az-AZ', { hour: '2-digit', minute: '2-digit' }),
        status: 'pending',
        alerted: false,
        items: table.orders.map(o => ({ name: o.name, quantity: o.quantity }))
    };

    kitchenOrders.push(newKitchenOrder);
    localStorage.setItem('rest_kitchen_orders', JSON.stringify(kitchenOrders));

    renderTables();

    alert(`${table.number} üçün sifariş uğurla mətbəxə göndərildi! 🍳`);
}

function closeAccount() {
    if (!currentTableId) return;
    const table = tables.find(t => t.id === currentTableId);
    let kitchenOrders = JSON.parse(localStorage.getItem('rest_kitchen_orders')) || [];

    const isStillCooking = kitchenOrders.some(ko => ko.tableName === table.number && ko.status === 'pending');

    if (isStillCooking) {
        alert(`⚠️ ÖDƏNİŞ BLOKLANDI!\nAşpaz hələ də ${table.number} üçün sifarişi hazır etməyib.\nYeməklər tam hazır olmadan hesabı bağlaya bilməzsiniz!`);
        return;
    }

    let tableTotal = table.orders.reduce((sum, item) => sum + (item.price * item.quantity), 0);

    const inputPaid = document.getElementById('amountPaidInput');
    const amountPaid = inputPaid ? (parseFloat(inputPaid.value) || 0) : 0;

    if (amountPaid < tableTotal) {
        alert("Ödənilən məbləğ ümumi cəmdən az ola bilməz!");
        return;
    }

    kitchenOrders = kitchenOrders.filter(ko => ko.tableName !== table.number);
    localStorage.setItem('rest_kitchen_orders', JSON.stringify(kitchenOrders));

    const archiveItem = {
        id: Math.floor(10000 + Math.random() * 90000),
        tableNumber: table.number,
        totalAmount: tableTotal,
        amountPaid: amountPaid,
        changeAmount: amountPaid - tableTotal,
        date: new Date().toLocaleString('az-AZ'),
        items: [...table.orders]
    };

    salesArchive.push(archiveItem);
    table.orders = [];
    table.status = 'empty';
    table.amountPaid = 0;

    updateStorage();
    calculateTotalEarnings();

    alert("Ödəniş uğurla tamamlandı və masa boşaldıldı!");

    currentTableId = null;

    const titleEl = document.getElementById('selectedTableTitle');
    if (titleEl) titleEl.innerText = "Masa seçilməyib";

    const pArea = document.getElementById('paymentArea');
    if (pArea) pArea.style.display = 'none';

    const btnCheck = document.getElementById('checkoutBtn');
    const btnRec = document.getElementById('receiptBtn');
    const btnKit = document.getElementById('kitchenBtn');

    if (btnCheck) btnCheck.disabled = true;
    if (btnRec) btnRec.disabled = true;
    if (btnKit) btnKit.disabled = true;

    renderTables();
    renderOrderDetails();
}

function previewReceipt() {
    if (!currentTableId) return;
    const table = tables.find(t => t.id === currentTableId);
    let tableTotal = table.orders.reduce((sum, item) => sum + (item.price * item.quantity), 0);

    const inputPaid = document.getElementById('amountPaidInput');
    const amountPaid = inputPaid ? (parseFloat(inputPaid.value) || 0) : 0;
    const change = amountPaid - tableTotal;

    const metaContainer = document.getElementById('receiptMeta');
    if (metaContainer) {
        metaContainer.innerHTML = `
                    <div class="receipt-line-item"><span>Tarix:</span> <span>${new Date().toLocaleDateString('az-AZ')}</span></div>
                    <div class="receipt-line-item"><span>Saat:</span> <span>${new Date().toLocaleTimeString('az-AZ', { hour: '2-digit', minute: '2-digit' })}</span></div>
                    <div class="receipt-line-item"><span>Masa:</span> <span><b>${table.number}</b></span></div>
                `;
    }

    const bodyContainer = document.getElementById('receiptItemsBody');
    if (bodyContainer) {
        bodyContainer.innerHTML = '';
        table.orders.forEach(item => {
            const itemTotal = item.price * item.quantity;
            const row = document.createElement('div');
            row.className = 'receipt-line-item';
            row.innerHTML = `<span>${item.name} x ${item.quantity}</span> <span>${itemTotal.toFixed(2)} AZN</span>`;
            bodyContainer.appendChild(row);
        });
    }

    const totalValEl = document.getElementById('receiptTotalVal');
    if (totalValEl) totalValEl.innerText = `${tableTotal.toFixed(2)} AZN`;

    const paymentDetails = document.getElementById('receiptPaymentDetails');
    if (paymentDetails) {
        paymentDetails.innerHTML = `
                    <div style="width: 100%;">
                        <div class="receipt-line-item" style="margin-top: 5px;"><span>Ödənilən:</span> <span>${amountPaid.toFixed(2)} AZN</span></div>
                        <div class="receipt-line-item"><span>Qalıq:</span> <span>${change >= 0 ? change.toFixed(2) : '0.00'} AZN</span></div>
                    </div>
                `;
    }

    const modalEl = document.getElementById('receiptModal');
    if (modalEl) modalEl.style.display = 'flex';
}

function resetSalesArchive() {
    if (confirm("Diqqət! Bütün satış tarixçəsi silinəcək və Excel hesabatı sıfırlanacaq. Bu əməliyyatı geri qaytarmaq olmaz! Əminsiniz?")) {
        localStorage.removeItem('rest_sales_archive');
        salesArchive = [];
        calculateTotalEarnings();
        alert("Satış hesabatı uğurla sıfırlandı! Yeni gündən başlaya bilərsiniz.");
    }
}

function closeReceiptModal() {
    const modalEl = document.getElementById('receiptModal');
    if (modalEl) modalEl.style.display = 'none';
}

function printReceipt() { window.print(); }

window.onload = function () {
    renderTables();
    renderMenu();
    calculateTotalEarnings();

    setInterval(() => {
        if (currentTableId) {
            renderOrderDetails();
        }
        renderTables();
    }, 2000);
};