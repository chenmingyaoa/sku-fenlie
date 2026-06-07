const form = document.getElementById("analyze-form");
const submitButton = document.getElementById("submit-btn");
const turnoverFileInput = document.getElementById("turnover-file");
const groupFilterList = document.getElementById("group-filter-list");
const siteFilter = document.getElementById("site-filter");

function renderGroupOptions(values, selectedValues = ["__all__"]) {
  if (!groupFilterList) return;
  const activeValues = new Set(selectedValues);
  groupFilterList.innerHTML = "";

  values.forEach((value) => {
    const label = document.createElement("label");
    label.className = "check-chip";

    const input = document.createElement("input");
    input.type = "checkbox";
    input.name = "group_filters";
    input.value = value;
    input.checked = activeValues.has(value);

    const text = document.createElement("span");
    text.textContent = value;

    label.appendChild(input);
    label.appendChild(text);
    groupFilterList.appendChild(label);
  });

  wireGroupCheckboxes();
}

function fillSiteSelect(values, selectedValue = "__all__") {
  if (!siteFilter) return;
  siteFilter.innerHTML = "";

  const defaultOption = document.createElement("option");
  defaultOption.value = "__all__";
  defaultOption.textContent = "全部站点";
  siteFilter.appendChild(defaultOption);

  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    option.selected = selectedValue === value;
    siteFilter.appendChild(option);
  });

  if (!values.includes(selectedValue)) {
    siteFilter.value = "__all__";
  }
}

function wireGroupCheckboxes() {
  const allCheckbox = document.querySelector('input[name="group_filters"][value="__all__"]');
  const groupCheckboxes = [...document.querySelectorAll('#group-filter-list input[name="group_filters"]')];
  if (!allCheckbox) return;

  allCheckbox.onchange = () => {
    if (allCheckbox.checked) {
      groupCheckboxes.forEach((checkbox) => {
        checkbox.checked = false;
      });
    }
  };

  groupCheckboxes.forEach((checkbox) => {
    checkbox.onchange = () => {
      if (checkbox.checked && allCheckbox.checked) {
        allCheckbox.checked = false;
      }
      if (!groupCheckboxes.some((item) => item.checked)) {
        allCheckbox.checked = true;
      }
    };
  });
}

async function loadFilterOptions(file) {
  if (!file) return;
  const data = new FormData();
  data.append("turnover_file", file);

  try {
    renderGroupOptions([], ["__all__"]);
    fillSiteSelect([], "__all__");

    const response = await fetch("/filter-options", {
      method: "POST",
      body: data,
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "读取筛选条件失败");
    }

    renderGroupOptions(payload.groups || [], ["__all__"]);
    fillSiteSelect(payload.sites || [], "__all__");
  } catch (error) {
    renderGroupOptions([], ["__all__"]);
    fillSiteSelect([], "__all__");
    window.alert(error.message);
  }
}

if (turnoverFileInput) {
  turnoverFileInput.addEventListener("change", (event) => {
    const file = event.target.files && event.target.files[0];
    loadFilterOptions(file);
  });
}

wireGroupCheckboxes();

if (form && submitButton) {
  form.addEventListener("submit", () => {
    submitButton.disabled = true;
    submitButton.textContent = "分析中，请稍候...";
  });
}
