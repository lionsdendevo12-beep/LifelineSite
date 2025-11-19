
const buttonsDiv = document.getElementById("buttons");
const gridDiv = document.getElementById("grid");
const modal = document.getElementById("modal");
const modalTitle = document.getElementById("modalTitle");
const modalType = document.getElementById("modalType");
const modalDescription = document.getElementById("modalDescription");
const modalImage = document.getElementById("modalImage");
const visitBtn = document.getElementById("visitBtn");
const closeModal = document.getElementById("closeModal");
const prevBtn = document.getElementById("prevBtn");
const nextBtn = document.getElementById("nextBtn");

let allData = [];
let currentType = null;
let currentIndex = -1;
let currentList = [];

function renderButtons(types) {
    buttonsDiv.innerHTML = "";
    if (types.length === 0) {
        buttonsDiv.innerHTML = "<p><em>No data loaded. Check your XLSX file.</em></p>";
        return;
    }

    const allBtn = document.createElement("button");
    allBtn.textContent = "All";
    allBtn.classList.add("active");
    allBtn.onclick = () => { currentType = null; renderGrid(); setActive(allBtn); };
    buttonsDiv.appendChild(allBtn);

    types.forEach(type => {
        const btn = document.createElement("button");
        btn.textContent = type;
        btn.onclick = () => { currentType = type; renderGrid(); setActive(btn); };
        buttonsDiv.appendChild(btn);
    });
}

function setActive(activeBtn) {
    document.querySelectorAll(".buttons button").forEach(btn => btn.classList.remove("active"));
    activeBtn.classList.add("active");
}

function renderGrid() {
    gridDiv.innerHTML = "";
    currentList = currentType ? allData.filter(d => d.type === currentType) : allData;

    currentList.forEach((item, i) => {
        const div = document.createElement("div");
        div.className = "card";

        div.innerHTML = `
            ${item.image ? `<img src="${item.image}" style="border:2px solid red; max-width:150px;">` : "<p>No img</p>"}
            <h2>${item.name || "Untitled"}</h2>
            <p>${item.type || "Unknown"}</p>
        `;
        console.log("Rendering item:", item.name, item.image?.substring(0, 40));
        div.onclick = () => openModal(i);
        gridDiv.appendChild(div);
    });
}

function openModal(index) {
    if (index < 0 || index >= currentList.length) return;

    currentIndex = index;
    const item = currentList[currentIndex];

    modal.style.display = "flex";
    modalTitle.textContent = item.name || "Untitled";
    modalType.textContent = `Type: ${item.type || "Unknown"}`;
    modalDescription.textContent = item.description || "No description available.";

    modalImage.src = item.image || "";
    modalImage.style.display = item.image ? "block" : "none";

    visitBtn.onclick = () => {
        if (item.website) window.open(item.website, "_blank");
    };
}

function closeModalFunc() {
    modal.style.display = "none";
    currentIndex = -1;
}

modal.onclick = (e) => { if (e.target === modal) closeModalFunc(); };
closeModal.onclick = closeModalFunc;

prevBtn.onclick = () => {
    if (currentIndex > 0) openModal(currentIndex - 1);
};

nextBtn.onclick = () => {
    if (currentIndex < currentList.length - 1) openModal(currentIndex + 1);
};

async function loadJSON() {
    try {
        const response = await fetch("data.json");
        const json = await response.json();

        allData = json.map(row => ({
            name: row.name || "",
            type: row.type || "",
            website: row.website || "",
            description: row.description || "",
            image: row.image || null
        }));


        const types = [...new Set(allData.map(r => r.type).filter(Boolean))];

        renderButtons(types);
        renderGrid();

    } catch (err) {
        console.error("Error loading data.json:", err);
        buttonsDiv.innerHTML = "<p style='color:red;'>Could not load data.json</p>";
    }
}

loadJSON();

