import React, { Fragment, useState, useEffect, createRef } from "react";
import Header from "./Header";
import FilePicker from "./FilePicker";
import Progress from "./Progress";
import Container from "./styles/Container";
import styled, { createGlobalStyle } from "styled-components";
import Preview from "./Preview";
/* global Button, console, Excel, Header, HeroList, HeroListItem, Progress */

const GlobalStyle = createGlobalStyle`
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

body {
  width: 100vw;
  height: 100vh;
  font-family: Arial, Helvetica, sans-serif;
}

#container {
  display: flex;
  min-height: 100vh;
  justify-content: center;
}
`;

const HiddenCanvas = styled.canvas`
  display: none;
`;

const sleep = ms => {
  return new Promise(resolve => setTimeout(resolve, ms));
};

const App = props => {
  // OfficeExtension.config.extendedErrorLogging = true;
  const { title, isOfficeInitialized } = props;
  const [image, setImage] = useState(null);
  const [imageData, setImageData] = useState(null);
  const [parsing, setParsing] = useState(false);
  const [painting, setPainting] = useState(false);
  const initialDimensions = {
    width: 10000,
    height: 10000
  };
  const [canvasDimensions, setCanvasDimensions] = useState(initialDimensions);
  const [outputSheet, setOutputSheet] = useState("");
  const canvas = createRef();

  const rgbToHex = (r, g, b) =>
    "#" +
    [r, g, b]
      .map(x => {
        const hex = x.toString(16);
        return hex.length === 1 ? "0" + hex : hex;
      })
      .join("");

  const afterImageChange = () => {
    const ctx = canvas.current.getContext("2d");
    const img = new Image();
    img.onload = () => {
      const width = img.width;
      const height = img.height;

      setCanvasDimensions({ width, height });
      ctx.drawImage(img, 0, 0, width, height);

      // URL.revokeObjectURL(img.src);

      /* Read pixel data */
      const imageData = ctx.getImageData(0, 0, width, height);
      const data = imageData.data;
      // => [r,g,b,a,...]

      const pixels = [];
      for (let i = 0; i < data.length; i += 4) {
        pixels.push(
          rgbToHex(
            data[i],
            data[i + 1],
            data[i + 2]
            // data[i+3] == alpha
          )
        );
      }
      setImageData({ width: img.width, height: img.height, pixels: pixels });
      setParsing(false);
    };
    img.src = image.src;
  };

  const createOutputSheet = async () => {
    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */
        if (image && imageData) {
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
          const outputName = createSheetName(sheetName);
          sheets.add(outputName);
          setOutputSheet(outputName);
        }
      });
    } catch (error) {
      console.error(error);
      setPainting(false);
    }
  };

  const formatOutputSheet = async () => {
    try {
      await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem(outputSheet);
        sheet.activate();
        const range = sheet.getRangeByIndexes(0, 0, imageData.height, imageData.width);
        range.format.columnWidth = 0.72;
        range.format.rowHeight = 0.72;
        range.untrack();
      });
      await paint();
      setPainting(false);
    } catch (error) {
      console.error(error);
      setPainting(false);
    }
  };

  const paint = async () => {
    // const blockSize = Math.max(parseInt(2000 / imageData.width), 1); // ~1k cells at a time
    // for (let i = 0; i < imageData.height; i += blockSize) {
    const blockSize = Math.max(parseInt(2000 / imageData.width), 1); // ~2k cells at a time
    for (let i = 0; i < imageData.height; i += blockSize) {
      try {
        await paintBlock(i, blockSize);
        await sleep(0);
      } catch (error) {
        console.log("Error painting block");
        console.error(error);
        break;
      }
    }
  };

  const paintBlock = async (startRow, blockSize) => {
    await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getItem(outputSheet);
      // Update the fill color
      const fillCell = (row, col) => {
        const cell = sheet.getCell(row, col);
        const color = imageData.pixels[col + row * imageData.width];
        cell.format.fill.color = color;
        // call untrack() to release the range from memory
        cell.untrack();
      };
      const fillRow = row => {
        for (let j = 0; j < imageData.width; j++) {
          fillCell(row, j);
        }
      };
      const render = () => {
        console.log(`${startRow} to ${Math.min(startRow + blockSize, imageData.height)}`);
        for (let i = startRow; i < Math.min(startRow + blockSize, imageData.height); i++) {
          fillRow(i);
        }
      };
      render();
    });
  };

  const imageChange = e => {
    const blob = e.target.files[0];
    const removeExtension = x => x.substr(0, x.lastIndexOf(".")) || x;
    setImage({
      displayName: removeExtension(blob.name),
      fileName: blob.name,
      src: URL.createObjectURL(blob)
    });
  };

  const click = () => {
    setPainting(true);
  };

  useEffect(() => {
    if (painting) {
      formatOutputSheet();
    }
  }, [outputSheet]);

  useEffect(() => {
    if (painting) {
      createOutputSheet();
    }
  }, [painting]);

  useEffect(() => {
    if (parsing) {
      afterImageChange();
    }
  }, [parsing]);

  useEffect(() => {
    if (image) {
      setParsing(true);
    }
  }, [image]);

  if (!isOfficeInitialized) {
    return (
      <Fragment>
        <GlobalStyle />
        <Progress
          title={title}
          logo="assets/logo-filled.png"
          message={parsing ? "Parsing Image" : painting ? "Painting Image" : "Loading"}
        />
      </Fragment>
    );
  }

  return (
    <Fragment>
      <GlobalStyle />
      <Container>
        <HiddenCanvas width={canvasDimensions.width} height={canvasDimensions.height} ref={canvas} />
        <Header logo="assets/logo-filled.png" title={title} message="Welcome" />
        <FilePicker handleChange={imageChange} />
        {image && <Preview image={image} handleClick={click} busy={parsing || painting} />}
      </Container>
    </Fragment>
  );
};

export default App;
