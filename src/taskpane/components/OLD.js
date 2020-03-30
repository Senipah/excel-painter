const paint = async () => {
  try {
    await Excel.run(async context => {
      /**
       * Insert your Excel code here
       */
      if (image && imageData) {
        const app = context.workbook.application;
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
        const existingSheetNames = sheets.items.map(e => e.name);
        const NAME_LENGTH = 26; // max length is 31. deduct 5 in case suffix required
        const sheetName = image.displayName.replace(/\\\/\*\?:\[]/gi, "").substring(0, NAME_LENGTH);
        const createSheetName = name => {
          let exists = false;
          let ctr = 0;
          let suffix = "";
          do {
            exists = existingSheetNames.includes(name + suffix);
            if (exists) {
              ctr += 1;
              suffix = ` (${ctr})`;
            }
          } while (exists === true);
          return name + suffix;
        };
        const outputSheet = sheets.add(createSheetName(sheetName));
        outputSheet.activate();
        app.suspendApiCalculationUntilNextSync();
        app.suspendScreenUpdatingUntilNextSync();
        const range = outputSheet.getRangeByIndexes(0, 0, imageData.height, imageData.width);
        range.load("address");
        await context.sync();
        range.format.columnWidth = 0.72;
        range.format.rowHeight = 0.72;
        range.untrack();
        // Update the fill color
        const fillCell = (row, col) => {
          const cell = outputSheet.getCell(row, col);
          const color = imageData.pixels[col + row * imageData.width];
          cell.format.fill.color = color;
          // cell.format.fill.color = "#000000";
          // call untrack() to release the range from memory
          cell.untrack();
        };
        const fillRow = async row => {
          for (let j = 0; j < imageData.width; j++) {
            fillCell(row, j);
            if (j % 100 === 0 && j > 0) {
              await sleep(0);
              await context.sync();
            }
          }
        };
        const render = async () => {
          for (let i = 0; i < imageData.height; i++) {
            await fillRow(i);
            if (i % 10 === 0 && i > 0) {
              console.log(`Finished row: ${i}`);
            }
          }
        };
        OfficeExtension.config.extendedErrorLogging = true;
        await render();
        // console.log("Finished render. syncing...");
        // await sleep(1000);
        await context.sync();
        setPainting(false);
      }
    });
  } catch (error) {
    console.error(error);
    setPainting(false);
  }
};
