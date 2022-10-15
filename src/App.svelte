<script>
  // @ts-nocheck
  let studentName = "Радюк Максим Александрович";
  let groupName = "22ИТ-1";
  let teacherName = "Вишняков Ю.М.";
  let labNumber = "1";
  let labName = "Целочисленные типы";

  function generate() {
    const doc = new docx.Document({
      styles: {
        paragraphStyles: [
          {
            id: "headingPara",
            name: "Heading Para",
            basedOn: "Normal",
            quickFormat: true,
            run: {
              size: 28,
            },
          },
        ],
      },
      sections: [
        {
          properties: {},
          children: [
            new docx.Paragraph({
              children: [
                new docx.TextRun({
                  style: "headingPara",
                  text: "Студент " + studentName,
                }),
              ],
            }),
            new docx.Paragraph({
              children: [
                new docx.TextRun({
                  style: "headingPara",
                  text: "Группа " + groupName,
                }),
              ],
            }),
            new docx.Paragraph({
              children: [
                new docx.TextRun({
                  style: "headingPara",
                  text: "Преподаватель " + teacherName,
                }),
              ],
            }),
            new docx.Paragraph({
              children: [
                new docx.TextRun({
                  style: "headingPara",
                  text: "",
                }),
              ],
            }),
            new docx.Paragraph({
              children: [
                new docx.TextRun({
                  style: "headingPara",
                  text: "РАБОТА #" + labNumber + " - " + labName.toUpperCase(),
                }),
              ],
            }),
          ],
        },
      ],
    });

    docx.Packer.toBlob(doc).then((blob) => {
      console.log(blob);
      saveAs(blob, "example.docx");
      console.log("Document created successfully");
    });
  }
</script>

<main>
  <h2>Lab generator</h2>
  <input bind:value={studentName} placeholder="Имя студента" />
  <input bind:value={groupName} placeholder="Название группы" />
  <input bind:value={teacherName} placeholder="Имя преподавателя" />
  <input bind:value={labNumber} placeholder="Номер лабаратнорной" />
  <input bind:value={labName} placeholder="Название лабараторной" />
  <button on:click={generate}>Generate</button>
</main>

<style>
  main {
    display: flex;
    flex-direction: column;
  }

  input, button {
    width: 250px;
  }

</style>
