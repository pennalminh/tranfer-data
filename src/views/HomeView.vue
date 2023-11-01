<template>
  <div>
    <input id="input-file" type="file" accept=".xlsx" />
    <select v-model="sheet_name">
      <option v-for="name in listSheet" :key="name" :value="name">
        {{ name }}
      </option>
    </select>
  </div>
  <br />
  <div>
    <button @click="handleData()">Load dữ liệu</button>
  </div>
</template>

<script setup lang="ts">
// Nhập file
import readXlsxFile from "read-excel-file";
import { readSheetNames } from "read-excel-file";
import { onMounted, ref } from "vue";

const listSheet = ref([]);
const sheet_name = ref("");
const data_excel = ref([]);
let input = null;

onMounted(() => {
  // File.
  input = document.getElementById("input-file");
  input.addEventListener("change", () => {
    readSheetNames(input.files[0]).then((sheetNames) => {
      listSheet.value = sheetNames;
    });
  });
});

const handleData = () => {
  let content = "";
  readXlsxFile(input.files[0], { sheet: sheet_name.value }).then((rows) => {
    data_excel.value = rows;
    let count = 0;
    for (let index = 0; index < rows.length; index++) {
      if (typeof rows[index][3] == "number") {
        count++;
        if (count % 2 == 0) {
          if (Math.abs(rows[index][3] - rows[index - 1][3]) > 5) {
            let sum = rows[index][3] + rows[index - 1][3];
            let ran = (Math.random() * 2.5 + 1) / 2;

            content += sum / 2 + ran + "\n" + (sum / 2 - ran) + "\n";
          } else {
            content += rows[index - 1][3] + "\n" + rows[index][3] + "\n";
          }
        }
      }
    }

    // Create element with <a> tag
    const link = document.createElement("a");

    // Create a blog object with the file content which you want to add to the file
    const file = new Blob([content], { type: "text/plain" });

    // Add file content in the object URL
    link.href = URL.createObjectURL(file);

    // Add file name
    link.download = "data.txt";

    link.click();
  });
};
</script>
