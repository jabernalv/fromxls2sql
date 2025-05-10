const fileInput = document.getElementById("fileInput");
const processBtn = document.getElementById("processBtn");
const output = document.getElementById("output");
const resultBox = document.getElementById("resultContainer");
const fileAlert = document.getElementById("fileAlert");
const tableNameInput = document.getElementById("tableName");
const includeCreateInput = document.getElementById("includeCreate");
const includeMigrationInput = document.getElementById("includeMigration");
const tableNameFeedback = document.getElementById("tableNameFeedback");
const downloadBtn = document.getElementById("downloadBtn");
const dialect = document.getElementById("sqlDialect");

let selectedFile = null;

tableNameInput.addEventListener("input", () => {
  const val = tableNameInput.value.trim().toLowerCase();
  tableNameInput.value = val;

  const valid = /^[a-z][a-z0-9_]*$/.test(val);
  if (val === "") {
    tableNameFeedback.classList.remove("text-green-600", "text-red-600");
    tableNameFeedback.classList.add("hidden");
  } else if (!valid) {
    tableNameFeedback.textContent = "❌ Nombre inválido";
    tableNameFeedback.classList.remove("hidden", "text-green-600");
    tableNameFeedback.classList.add("text-red-600");
  } else {
    tableNameFeedback.textContent = "✅ Nombre válido";
    tableNameFeedback.classList.remove("hidden", "text-red-600");
    tableNameFeedback.classList.add("text-green-600");
  }

  updateProcessButtonState();
});

function updateProcessButtonState() {
  const validTableName = /^[a-z][a-z0-9_]*$/.test(tableNameInput.value.trim());
  const enable = !!selectedFile && validTableName;

  processBtn.disabled = !enable;
  processBtn.classList.toggle("bg-blue-600", enable);
  processBtn.classList.toggle("bg-blue-300", !enable);
  processBtn.classList.toggle("cursor-not-allowed", !enable);
}

fileInput.addEventListener("change", function (event) {
  selectedFile = event.target.files[0];
  updateProcessButtonState();

  const fileLabel = document.getElementById("fileLabel");

  if (selectedFile) {
    fileLabel.innerHTML =
      '<i data-lucide="check" class="w-4 h-4 text-green-600"></i> Archivo seleccionado';
  } else {
    fileLabel.innerHTML =
      '<i data-lucide="upload" class="w-4 h-4"></i> Elegir archivo';
  }

  lucide.createIcons(); // re-render íconos
});

processBtn.addEventListener("click", function () {
  fileAlert.classList.add("hidden");
  resultBox.classList.add("hidden");
  output.value = "";

  if (!selectedFile) return;

  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      if (workbook.SheetNames.length !== 1) {
        showError("El archivo debe contener exactamente una hoja.");
        return;
      }

      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      if (rows.length === 0 || rows[0].length === 0) {
        showError("No se detectaron encabezados en la primera fila.");
        return;
      }

      const headers = rows[0].map((h) => h.toLowerCase());
      const fieldPattern = /^[a-z][a-z0-9_]*$/;
      const invalid = headers.filter((h) => !fieldPattern.test(h));

      if (invalid.length > 0) {
        showError("Nombres de columnas inválidos: " + invalid.join(", "));
        return;
      }

      const columnTypes = inferColumnTypes(rows);
      const tableName = tableNameInput.value.trim();
      const includeCreate = includeCreateInput.value === "yes";
      const includeMigration = includeMigrationInput?.value === "yes";

      let sql = "";
      if (includeCreate) {
        sql += generateCreateTableSQL(tableName, headers, columnTypes) + "\n";
      }
      sql += generateInsertStatements(tableName, headers, rows);

      if (includeMigration) {
        sql +=
          "\n\n/* === Yii2 Migration === */\n\n" +
          generateYiiMigration(tableName, headers, columnTypes, rows);
      }

      output.value = sql;
      resultBox.classList.remove("hidden");
    } catch (err) {
      showError("Error al procesar el archivo: " + err.message);
    }
  };

  reader.readAsArrayBuffer(selectedFile);
});

function showError(message) {
  fileAlert.textContent = message;
  fileAlert.classList.remove("hidden");
}

function inferColumnTypes(rows) {
  const types = [];
  const maxSamples = 1000;
  const isInt = (val) => /^-?\d+$/.test(val);
  const isFloat = (val) => /^-?\d+\.\d+$/.test(val);
  const isDate = (val) => /^\d{4}-\d{2}-\d{2}$/.test(val);

  const columnCount = rows[0].length;
  for (let col = 0; col < columnCount; col++) {
    const values = [];
    for (let i = 1; i < rows.length && values.length < maxSamples; i++) {
      const cell = rows[i][col];
      if (cell !== undefined && cell !== null && String(cell).trim() !== "") {
        values.push(String(cell));
      }
    }

    let allInt = true,
      allFloat = true,
      allDate = true,
      maxLength = 0;

    for (const val of values) {
      if (!isInt(val)) allInt = false;
      if (!isFloat(val) && !isInt(val)) allFloat = false;
      if (!isDate(val)) allDate = false;
      maxLength = Math.max(maxLength, val.length);
    }
    maxLength = maxLength * 10;

    if (allInt) {
      types.push("INT");
    } else if (allFloat) {
      types.push("FLOAT");
    } else if (allDate) {
      types.push("DATE");
    } else {
      types.push(`VARCHAR(${Math.min(maxLength, 255)})`);
    }
  }

  return types;
}

function generateCreateTableSQL(tableName, headers, types) {
  const selected = dialect.value;

  let wrapStart = "`",
    wrapEnd = "`",
    tableWrap = "`" + tableName + "`";
  if (selected === "postgresql" || selected === "oracle") {
    wrapStart = wrapEnd = '"';
    tableWrap = `"${tableName}"`;
  } else if (selected === "sqlserver") {
    wrapStart = "[";
    wrapEnd = "]";
    tableWrap = `[${tableName}]`;
  }

  const cols = headers.map((name, i) => {
    let type = types[i];
    const isId = name === "id" && type === "INT";

    // Adaptar tipos y autoincremento por dialecto
    if (isId) {
      if (selected === "mysql") type = "INT AUTO_INCREMENT PRIMARY KEY";
      else if (selected === "postgresql") type = "SERIAL PRIMARY KEY";
      else if (selected === "sqlserver") type = "INT IDENTITY(1,1) PRIMARY KEY";
      else if (selected === "oracle")
        type = "NUMBER GENERATED BY DEFAULT ON NULL AS IDENTITY PRIMARY KEY";
    } else {
      if (selected === "oracle") {
        if (type.startsWith("VARCHAR")) {
          const size = type.match(/\\d+/)?.[0] || 255;
          type = `VARCHAR2(${size})`;
        } else if (type === "INT" || type === "FLOAT") {
          type = "NUMBER";
        }
      } else if (selected === "sqlserver") {
        if (type === "FLOAT") type = "REAL";
        if (type === "INT") type = "INT";
      }
    }

    return `  ${wrapStart}${name}${wrapEnd} ${type}`;
  });

  let ifNotExists =
    selected === "mysql" || selected === "postgresql" ? "IF NOT EXISTS " : "";

  return `CREATE TABLE ${ifNotExists}${tableWrap} (\n${cols.join(",\n")}\n);`;
}

function generateInsertStatements(tableName, headers, rows) {
  const selected = dialect.value;

  let wrapStart = "`",
    wrapEnd = "`";
  if (selected === "postgresql" || selected === "oracle") {
    wrapStart = wrapEnd = '"';
  } else if (selected === "sqlserver") {
    wrapStart = "[";
    wrapEnd = "]";
  }

  const tableWrap =
    selected === "sqlserver"
      ? `[${tableName}]`
      : `${wrapStart}${tableName}${wrapEnd}`;

  const inserts = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;

    const values = headers.map((_, j) => {
      const val = row[j];
      if (val === undefined || val === null || val === "") return "NULL";
      const raw = String(val);
      const needsQuote = isNaN(raw) || /^0\\d/.test(raw) || /[^\\d.]/.test(raw);
      return needsQuote ? `'${raw.replace(/'/g, "''")}'` : raw;
    });

    const columns = headers.map((h) => `${wrapStart}${h}${wrapEnd}`).join(", ");
    inserts.push(
      `INSERT INTO ${tableWrap} (${columns}) VALUES (${values.join(", ")});`
    );
  }

  return inserts.join("");
}

function generateYiiMigration(tableName, headers, types, rows) {
  const className = `m${new Date()
    .toISOString()
    .slice(0, 10)
    .replace(/-/g, "")}_create_${tableName}_table`;

  const fieldLines = headers.map((name, i) => {
    const type = types[i];
    if (name === "id" && type === "INT") {
      return `            '${name}' => $this->primaryKey(),`;
    } else if (type.startsWith("VARCHAR")) {
      const size = type.match(/\\d+/)?.[0] || 255;
      return `            '${name}' => $this->string(${size}),`;
    } else if (type === "INT") {
      return `            '${name}' => $this->integer(),`;
    } else if (type === "FLOAT") {
      return `            '${name}' => $this->float(),`;
    } else if (type === "DATE") {
      return `            '${name}' => $this->date(),`;
    } else {
      return `            '${name}' => $this->text(),`;
    }
  });

  const insertCommands = [];
  const columnList = headers.map((h) => `\`${h}\``).join(", ");

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;

    const values = headers.map((_, j) => {
      const val = row[j];
      if (val === undefined || val === null || val === "") return "NULL";
      const raw = String(val);
      const needsQuote = isNaN(raw) || /^0\\d/.test(raw) || /[^\\d.]/.test(raw);
      return needsQuote ? `'${raw.replace(/'/g, "''")}'` : raw;
    });

    insertCommands.push(
      `        Yii::$app->db->createCommand("INSERT INTO \\\`${tableName}\\\` (${columnList}) VALUES (${values.join(
        ", "
      )});")->execute();`
    );
  }

  return `<?php

use yii\\db\\Migration;

/**
 * Handles the creation of table \\\`{{%${tableName}}}\\\`.
 */
class ${className} extends Migration
{
    public function safeUp()
    {
        $this->createTable('{{%${tableName}}}', [
${fieldLines.join("\n")}
        ]);

${insertCommands.join("\n")}
    }

    public function safeDown()
    {
        $this->dropTable('{{%${tableName}}}');
    }
}
`;
}
downloadBtn.addEventListener("click", function () {
  const tableName = tableNameInput.value.trim() || "archivo";
  const blob = new Blob([output.value], { type: "text/sql;charset=utf-8;" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `${tableName}.sql`;
  link.click();
});
