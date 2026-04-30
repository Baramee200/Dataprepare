let previewRows = [];

document.getElementById("themeToggle").addEventListener("click", () => {
  document.body.classList.toggle("dark-mode");
});

function addStructure() {
  const officeName = prompt("กรอกชื่อสำนักงาน");
  if (!officeName) return;

  const container = document.getElementById("structures");
  const card = document.createElement("div");
  card.className = "card";
  card.dataset.index = document.querySelectorAll(".card").length + 1;

  const header = document.createElement("div");
  header.className = "card-header";
  header.innerHTML = `<h3>${officeName}</h3>
                      <span class="summary">ยังไม่มีข้อมูล</span>
                      <button class="delete" onclick="deleteStructure(this)">🗑 ลบสำนักงาน</button>`;
  header.onclick = (e) => {
    if (e.target.tagName !== "BUTTON") card.classList.toggle("open");
  };

  const body = document.createElement("div");
  body.className = "card-body";

  const ul = document.createElement("ul");
  ul.className = "structure";
  ul.dataset.index = card.dataset.index;

  const li = document.createElement("li");
  li.innerHTML = `
    <input class="officeName" value="${officeName}" placeholder="สำนัก" oninput="updateOfficeName(this)" />
    <label><input type="checkbox" class="skipRegister"> ไม่สร้างทะเบียน</label>
    <label><input type="checkbox" class="sendOnly"> สร้างเฉพาะทะเบียนส่ง</label>
    <label><input type="checkbox" class="receiveOnly"> สร้างเฉพาะทะเบียนรับ</label>
    <button class="toggleActions" onclick="toggleActions(this)">⚙ จัดการ</button>
    <div class="actions hidden">
      <button onclick="addNode(this.closest('li').querySelector('ul'),'กอง')">+ กอง</button>
      <button onclick="addNode(this.closest('li').querySelector('ul'),'แผนก')">+ กลุ่ม/ฝ่าย</button>
      <button onclick="addNode(this.closest('li').querySelector('ul'),'ส่วนงาน')">+ ส่วนงาน/แผนก</button>
      <button class="delete hidden" onclick="deleteNode(this)">🗑 ลบ</button>
    </div>
    <ul></ul>
  `;
  ul.appendChild(li);

  body.appendChild(ul);
  card.appendChild(header);
  card.appendChild(body);
  container.appendChild(card);
}

function deleteStructure(el) { el.closest(".card").remove(); }

function addNode(parent, type) {
  const li = document.createElement("li");
  li.innerHTML = `
    <span class="toggle" onclick="toggleNode(this)">▶</span>
    <input placeholder="${type}" />
    <label><input type="checkbox" class="skipRegister"> ไม่สร้างทะเบียน</label>
    <label><input type="checkbox" class="sendOnly"> สร้างเฉพาะทะเบียนส่ง</label>
    <label><input type="checkbox" class="receiveOnly"> สร้างเฉพาะทะเบียนรับ</label>
    <button class="toggleActions" onclick="toggleActions(this)">⚙ จัดการ</button>
    <div class="actions hidden">
      <button onclick="addNode(this.closest('li').querySelector('ul'),'กอง')">+ กอง</button>
      <button onclick="addNode(this.closest('li').querySelector('ul'),'แผนก')">+ แผนก</button>
      <button onclick="addNode(this.closest('li').querySelector('ul'),'ส่วนงาน')">+ ส่วนงาน</button>
      <!-- เอา hidden ออก -->
      <button class="delete" onclick="deleteNode(this)">🗑 ลบ</button>
    </div>
    <ul></ul>
  `;
  parent.appendChild(li);
}
  const parentDeleteBtn = parent.closest("li")?.querySelector(".delete");
  if (parentDeleteBtn) {
    parentDeleteBtn.classList.remove("hidden");
  }


function deleteNode(el) {
  const li = el.closest("li");
  li.remove();
}

  const parentUl = li.parentNode;
  if (parentUl && parentUl.children.length === 0) {
    const parentDeleteBtn = parentUl.closest("li")?.querySelector(".delete");
    if (parentDeleteBtn) {
      parentDeleteBtn.classList.add("hidden");
    }
  }


function toggleNode(el) {
  const li = el.parentNode;
  li.classList.toggle("collapsed");
  el.textContent = li.classList.contains("collapsed") ? "▶" : "▼";
}

function traverse(node, structureIndex) {
  const inputs = node.querySelectorAll(":scope > li");
  inputs.forEach(li => {
    const input = li.querySelector("input[type='text'], input:not([type])");
    const skip = li.querySelector(".skipRegister")?.checked;
    const sendOnly = li.querySelector(".sendOnly")?.checked;
    const receiveOnly = li.querySelector(".receiveOnly")?.checked; // เพิ่มตรงนี้
    const name = input?.value.trim();

    if (name) {
      if (skip) {
        previewRows.push([`ส่วนงาน ${structureIndex}`, name, "", structureIndex]);
      } else if (sendOnly) {
        previewRows.push([`ส่วนงาน ${structureIndex}`, name, `(ทะเบียนส่ง) ${name}`, structureIndex]);
      } else if (receiveOnly) {
        previewRows.push([`ส่วนงาน ${structureIndex}`, name, `(ทะเบียนรับ) ${name}`, structureIndex]);
      } else {
        previewRows.push([`ส่วนงาน ${structureIndex}`, name, `(ทะเบียนรับ) ${name}`, structureIndex]);
        previewRows.push([`ส่วนงาน ${structureIndex}`, name, `(ทะเบียนส่ง) ${name}`, structureIndex]);
      }
    }

    const childUl = li.querySelector("ul");
    if (childUl) traverse(childUl, structureIndex);
  });
}


function preview() {
  previewRows = [];
  const structures = document.querySelectorAll(".structure");
  structures.forEach((s, idx) => traverse(s, idx+1));

  let html = "<table><tr><th>ส่วนงาน</th><th>ชื่อหน่วยงาน</th><th>ทะเบียน</th></tr>";
  previewRows.forEach(row => {
    html += `<tr><td>${row[0]}</td><td>${row[1]}</td><td>${row[2]}</td></tr>`;
  });
  html += "</table>";
  document.getElementById("preview").innerHTML = html;
}

function exportExcel() {
  if (previewRows.length === 0) return alert("กรุณา Preview ก่อน Export");

  const ws = XLSX.utils.aoa_to_sheet([
    ["ส่วนงาน","ชื่อหน่วยงาน","ทะเบียน"],
    ...previewRows.map(r => [r[0], r[1], r[2]])
  ]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "ส่วนงาน");

  const now = new Date();
  const dateStr = now.toISOString().split("T")[0];
  XLSX.writeFile(wb, `โครงสร้างหน่วยงาน_${dateStr}.xlsx`);
}

function updateOfficeName(input) {
  const card = input.closest(".card");
  const headerTitle = card.querySelector(".card-header h3");
  const value = input.value.trim();
  headerTitle.textContent = value ? value : `สำนักงาน ${card.dataset.index}`;
}

function toggleActions(btn) {
  const li = btn.closest("li");
  const actions = li.querySelector(".actions");
  actions.classList.toggle("hidden");
}
