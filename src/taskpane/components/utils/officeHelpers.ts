// declare const Excel: any;
// export interface LinkMatch {
//   linkId: string;
//   matchedRange: string;
// }

// interface ExcelRect {
//   startRow: number;
//   endRow: number;
//   startCol: number;
//   endCol: number;
// }

// export const generateUUID = (): string => {
//   return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
//     const r = (Math.random() * 16) | 0;
//     const v = c === "x" ? r : (r & 0x3) | 0x8;
//     return v.toString(16);
//   });
// };

// export const getFileNameFromUrl = (url: string): string => {
//   if (!url || url === "excel-local-livelink") return "Unsaved Workbook.xlsx";
//   try {
//     const decodedUrl = decodeURIComponent(url);
//     const parts = decodedUrl.split(/[\\/]/);
//     return parts[parts.length - 1] || "Workbook.xlsx";
//   } catch (e) {
//     return "Workbook.xlsx";
//   }
// };

// const colLetterToNumber = (letter: string): number => {
//   let num = 0;
//   for (let i = 0; i < letter.length; i++) {
//     num = num * 26 + (letter.charCodeAt(i) - 64);
//   }
//   return num;
// };

// const parseRangeAddress = (address: string): ExcelRect => {
//   const cleanAddress = address.replace(/[$]/g, "");
//   const parts = cleanAddress.split(":");

//   const parseCell = (cell: string) => {
//     const match = cell.match(/^([A-Z]+)([0-9]+)$/i);
//     if (!match) return { row: 0, col: 0 };
//     return {
//       row: parseInt(match[2], 10),
//       col: colLetterToNumber(match[1].toUpperCase()),
//     };
//   };

//   const start = parseCell(parts[0]);
//   const end = parts[1] ? parseCell(parts[1]) : start;

//   return {
//     startRow: Math.min(start.row, end.row),
//     endRow: Math.max(start.row, end.row),
//     startCol: Math.min(start.col, end.col),
//     endCol: Math.max(start.col, end.col),
//   };
// };

// const rangesIntersect = (rect1: ExcelRect, rect2: ExcelRect): boolean => {
//   return !(
//     rect1.startRow > rect2.endRow ||
//     rect1.endRow < rect2.startRow ||
//     rect1.startCol > rect2.endCol ||
//     rect1.endCol < rect2.startCol
//   );
// };

// const isCoordinates = (str: string): boolean => {
//   return /^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$/i.test(str.replace(/[$]/g, ""));
// };

// export const getExistingLinkId = async (
//   sheetName: string,
//   rangeAddress: string
// ): Promise<LinkMatch | null> => {
//   return await Excel.run(async (context: any) => {
//     const parts = context.workbook.customXmlParts;
//     parts.load("items");
//     await context.sync();

//     const xmlBlobs = parts.items.map((part: any) => {
//       return {
//         id: part.id,
//         xmlBlob: part.getXml(),
//       };
//     });
//     await context.sync();

//     const parser = new DOMParser();
//     const cleanSheet = sheetName.trim().toLowerCase();
//     const cleanRange = rangeAddress.trim().toLowerCase();

//     for (let item of xmlBlobs) {
//       try {
//         const xmlText = item.xmlBlob.value;
//         if (xmlText && xmlText.includes("LiveLink")) {
//           const xmlDoc = parser.parseFromString(xmlText, "text/xml");
//           const savedSheet = (xmlDoc.getElementsByTagName("SheetName")[0]?.textContent || "")
//             .trim()
//             .toLowerCase();
//           const savedRange = (xmlDoc.getElementsByTagName("RangeAddress")[0]?.textContent || "")
//             .trim()
//             .toLowerCase();
//           const savedLinkId = xmlDoc.getElementsByTagName("LinkId")[0]?.textContent || "";

//           if (savedSheet === cleanSheet) {
//             if (savedRange === cleanRange) {
//               return {
//                 linkId: savedLinkId,
//                 matchedRange: savedRange,
//               };
//             }

//             if (isCoordinates(savedRange) && isCoordinates(cleanRange)) {
//               const savedRect = parseRangeAddress(savedRange);
//               const currentRect = parseRangeAddress(cleanRange);

//               if (rangesIntersect(currentRect, savedRect)) {
//                 return {
//                   linkId: savedLinkId,
//                   matchedRange: savedRange,
//                 };
//               }
//             }
//           }
//         }
//       } catch (partErr) {
//         console.error("Failed to parse single Custom XML Part safely:", partErr);
//       }
//     }

//     for (let item of xmlBlobs) {
//       try {
//         const xmlText = item.xmlBlob.value;
//         if (xmlText && xmlText.includes("LiveLink")) {
//           const xmlDoc = parser.parseFromString(xmlText, "text/xml");
//           const savedSheet = (xmlDoc.getElementsByTagName("SheetName")[0]?.textContent || "")
//             .trim()
//             .toLowerCase();
//           const savedRange = (xmlDoc.getElementsByTagName("RangeAddress")[0]?.textContent || "")
//             .trim()
//             .toLowerCase();
//           const savedLinkId = xmlDoc.getElementsByTagName("LinkId")[0]?.textContent || "";
//           const savedType = (xmlDoc.getElementsByTagName("Type")[0]?.textContent || "").trim();

//           if (savedSheet === cleanSheet && savedType === "Chart") {
//             return {
//               linkId: savedLinkId,
//               matchedRange: savedRange,
//             };
//           }
//         }
//       } catch (e) {
//         console.error("Failed to scan fallback chart links:", e);
//       }
//     }

//     return null;
//   });
// };

// export const getActiveSelection = async (
//   targetRange?: string
// ): Promise<{
//   isChart: boolean;
//   sheetName: string;
//   rangeAddress: string;
//   dataSnapshot: any;
//   chartTitle?: string;
// }> => {
//   return await Excel.run(async (context: any) => {
//     let isChart = false;
//     let sheetName = "";
//     let rangeAddress = "";
//     let dataSnapshot: any = null;

//     if (targetRange) {
//       const activeSheet = context.workbook.worksheets.getActiveWorksheet();
//       activeSheet.load("name");
//       await context.sync();

//       if (!isCoordinates(targetRange)) {
//         const charts = activeSheet.charts;
//         charts.load("items/name,items/title");
//         await context.sync();

//         const targetChart = charts.items.find(
//           (c: any) => c.name.toLowerCase() === targetRange.toLowerCase()
//         );
//         if (targetChart) {
//           const chartImage = targetChart.getImage();
//           targetChart.title.load("text");
//           await context.sync();

//           isChart = true;
//           sheetName = activeSheet.name;
//           rangeAddress = targetRange;
//           dataSnapshot = `data:image/png;base64,${chartImage.value}`;
//           return {
//             isChart,
//             sheetName,
//             rangeAddress,
//             dataSnapshot,
//             chartTitle: targetChart.title.text || targetChart.name,
//           };
//         }
//       }

//       const range = activeSheet.getRange(targetRange);
//       range.load("address");
//       await context.sync();

//       const rangeImage = range.getImage();
//       await context.sync();

//       sheetName = activeSheet.name;
//       rangeAddress = targetRange;
//       dataSnapshot = `data:image/png;base64,${rangeImage.value}`;
//       return { isChart, sheetName, rangeAddress, dataSnapshot };
//     }

//     try {
//       const activeChart = context.workbook.getSelectedChart();
//       activeChart.load(["name", "worksheet", "title"]);
//       activeChart.title.load("text");
//       await context.sync();

//       const chartImage = activeChart.getImage();
//       await context.sync();

//       isChart = true;
//       sheetName = activeChart.worksheet.name;
//       rangeAddress = activeChart.name;
//       dataSnapshot = `data:image/png;base64,${chartImage.value}`;
//       return {
//         isChart,
//         sheetName,
//         rangeAddress,
//         dataSnapshot,
//         chartTitle: activeChart.title.text || activeChart.name,
//       };
//     } catch (chartErr) {
//       console.log("[DEBUG] Active chart selection failed, falling back to Range check.");
//     }

//     try {
//       const activeRange = context.workbook.getSelectedRange();
//       activeRange.load(["address", "worksheet"]);
//       await context.sync();

//       const rangeImage = activeRange.getImage();
//       await context.sync();

//       sheetName = activeRange.worksheet.name;
//       const fullAddress = activeRange.address;
//       rangeAddress = fullAddress.includes("!") ? fullAddress.split("!")[1] : fullAddress;
//       rangeAddress = rangeAddress.replace(/['"]/g, "");

//       dataSnapshot = `data:image/png;base64,${rangeImage.value}`;
//       return { isChart, sheetName, rangeAddress, dataSnapshot };
//     } catch (rangeErr) {
//       try {
//         const activeSheet = context.workbook.worksheets.getActiveWorksheet();
//         activeSheet.load("name");
//         const charts = activeSheet.charts;
//         charts.load("items/name,items/title");
//         await context.sync();

//         if (charts.items.length > 0) {
//           const firstChart = charts.items[0];
//           const chartImage = firstChart.getImage();
//           firstChart.title.load("text");
//           await context.sync();

//           isChart = true;
//           sheetName = activeSheet.name;
//           rangeAddress = firstChart.name;
//           dataSnapshot = `data:image/png;base64,${chartImage.value}`;
//           return {
//             isChart,
//             sheetName,
//             rangeAddress,
//             dataSnapshot,
//             chartTitle: firstChart.title.text || firstChart.name,
//           };
//         }
//       } catch (innerErr) {}

//       throw new Error(
//         "No active range or chart detected. Please select a range or click on the chart container."
//       );
//     }
//   });
// };

// export const formatExcelRange = async (sheetName: string, rangeAddress: string) => {
//   await Excel.run(async (context: any) => {
//     const sheet = context.workbook.worksheets.getItem(sheetName);
//     const range = sheet.getRange(rangeAddress);

//     range.format.fill.color = "#F3F2F1";

//     const bottomBorder = range.format.borders.getItem("EdgeBottom");
//     bottomBorder.color = "#0078d4";
//     bottomBorder.style = "Continuous";
//     bottomBorder.weight = "Medium";

//     await context.sync();
//   });
// };

// export const clearExcelRangeFormat = async (sheetName: string, rangeAddress: string) => {
//   await Excel.run(async (context: any) => {
//     try {
//       const sheet = context.workbook.worksheets.getItem(sheetName);
//       const range = sheet.getRange(rangeAddress);
//       range.format.fill.clear();

//       const bottomBorder = range.format.borders.getItem("EdgeBottom");
//       bottomBorder.style = "None";

//       await context.sync();
//     } catch (e) {}
//   });
// };

// export const saveMetadataToCustomXml = async (
//   fileId: string,
//   linkId: string,
//   sheetName: string,
//   rangeAddress: string,
//   type: "Table" | "Chart"
// ) => {
//   const xmlString = `<LiveLink><LinkId>${linkId}</LinkId><FileId>${fileId}</FileId><SheetName>${sheetName}</SheetName><RangeAddress>${rangeAddress}</RangeAddress><Type>${type}</Type></LiveLink>`;

//   await Excel.run(async (context: any) => {
//     context.workbook.customXmlParts.add(xmlString);
//     await context.sync();
//     console.log(`[DEBUG] Successfully wrote Custom XML for Link ID: ${linkId}`);
//   });
// };

////////////////////////////////////


// declare const Excel: any;

// export interface LinkMatch {
//   linkId: string;
//   matchedRange: string;
// }

// interface ExcelRect {
//   startRow: number;
//   endRow: number;
//   startCol: number;
//   endCol: number;
// }

// export const generateUUID = (): string => {
//   return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
//     const r = (Math.random() * 16) | 0;
//     const v = c === "x" ? r : (r & 0x3) | 0x8;
//     return v.toString(16);
//   });
// };

// export const getFileNameFromUrl = (url: string): string => {
//   if (!url || url === "excel-local-livelink") return "Unsaved Workbook.xlsx";
//   try {
//     const decodedUrl = decodeURIComponent(url);
//     const parts = decodedUrl.split(/[\\/]/);
//     return parts[parts.length - 1] || "Workbook.xlsx";
//   } catch (e) {
//     return "Workbook.xlsx";
//   }
// };

// const colLetterToNumber = (letter: string): number => {
//   let num = 0;
//   for (let i = 0; i < letter.length; i++) {
//     num = num * 26 + (letter.charCodeAt(i) - 64);
//   }
//   return num;
// };

// const parseRangeAddress = (address: string): ExcelRect => {
//   const cleanAddress = address.replace(/[$]/g, "");
//   const parts = cleanAddress.split(":");

//   const parseCell = (cell: string) => {
//     const match = cell.match(/^([A-Z]+)([0-9]+)$/i);
//     if (!match) return { row: 0, col: 0 };
//     return {
//       row: parseInt(match[2], 10),
//       col: colLetterToNumber(match[1].toUpperCase()),
//     };
//   };

//   const start = parseCell(parts[0]);
//   const end = parts[1] ? parseCell(parts[1]) : start;

//   return {
//     startRow: Math.min(start.row, end.row),
//     endRow: Math.max(start.row, end.row),
//     startCol: Math.min(start.col, end.col),
//     endCol: Math.max(start.col, end.col),
//   };
// };

// const rangesIntersect = (rect1: ExcelRect, rect2: ExcelRect): boolean => {
//   return !(
//     rect1.startRow > rect2.endRow ||
//     rect1.endRow < rect2.startRow ||
//     rect1.startCol > rect2.endCol ||
//     rect1.endCol < rect2.startCol
//   );
// };

// const isCoordinates = (str: string): boolean => {
//   return /^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$/i.test(str.replace(/[$]/g, ""));
// };

// export const getExistingLinkId = async (
//   sheetName: string,
//   rangeAddress: string
// ): Promise<LinkMatch | null> => {
//   return await Excel.run(async (context: any) => {
//     const parts = context.workbook.customXmlParts;
//     parts.load("items");
//     await context.sync();

//     const xmlBlobs = parts.items.map((part: any) => {
//       return {
//         id: part.id,
//         xmlBlob: part.getXml(),
//       };
//     });
//     await context.sync();

//     const parser = new DOMParser();
//     const cleanSheet = sheetName.trim().toLowerCase();
//     const cleanRange = rangeAddress.trim().toLowerCase();

//     for (let item of xmlBlobs) {
//       try {
//         const xmlText = item.xmlBlob.value;
//         if (xmlText && xmlText.includes("LiveLink")) {
//           const xmlDoc = parser.parseFromString(xmlText, "text/xml");
//           const savedSheet = (xmlDoc.getElementsByTagName("SheetName")[0]?.textContent || "")
//             .trim()
//             .toLowerCase();
//           const savedRange = (xmlDoc.getElementsByTagName("RangeAddress")[0]?.textContent || "")
//             .trim()
//             .toLowerCase();
//           const savedLinkId = xmlDoc.getElementsByTagName("LinkId")[0]?.textContent || "";

//           if (savedSheet === cleanSheet) {
//             if (savedRange === cleanRange) {
//               return {
//                 linkId: savedLinkId,
//                 matchedRange: savedRange,
//               };
//             }

//             if (isCoordinates(savedRange) && isCoordinates(cleanRange)) {
//               const savedRect = parseRangeAddress(savedRange);
//               const currentRect = parseRangeAddress(cleanRange);

//               if (rangesIntersect(currentRect, savedRect)) {
//                 return {
//                   linkId: savedLinkId,
//                   matchedRange: savedRange,
//                 };
//               }
//             }
//           }
//         }
//       } catch (partErr) {
//         console.error("Failed to parse single Custom XML Part safely:", partErr);
//       }
//     }

//     for (let item of xmlBlobs) {
//       try {
//         const xmlText = item.xmlBlob.value;
//         if (xmlText && xmlText.includes("LiveLink")) {
//           const xmlDoc = parser.parseFromString(xmlText, "text/xml");
//           const savedSheet = (xmlDoc.getElementsByTagName("SheetName")[0]?.textContent || "")
//             .trim()
//             .toLowerCase();
//           const savedRange = (xmlDoc.getElementsByTagName("RangeAddress")[0]?.textContent || "")
//             .trim()
//             .toLowerCase();
//           const savedLinkId = xmlDoc.getElementsByTagName("LinkId")[0]?.textContent || "";
//           const savedType = (xmlDoc.getElementsByTagName("Type")[0]?.textContent || "").trim();

//           if (savedSheet === cleanSheet && savedType === "Chart") {
//             return {
//               linkId: savedLinkId,
//               matchedRange: savedRange,
//             };
//           }
//         }
//       } catch (e) {
//         console.error("Failed to scan fallback chart links:", e);
//       }
//     }

//     return null;
//   });
// };

// export const getActiveSelection = async (
//   targetRange?: string
// ): Promise<{
//   isChart: boolean;
//   sheetName: string;
//   rangeAddress: string;
//   dataSnapshot: any;
//   chartTitle?: string;
// }> => {
//   return await Excel.run(async (context: any) => {
//     let isChart = false;
//     let sheetName = "";
//     let rangeAddress = "";
//     let dataSnapshot: any = null;

//     if (targetRange) {
//       const activeSheet = context.workbook.worksheets.getActiveWorksheet();
//       activeSheet.load("name");
//       await context.sync();

//       if (!isCoordinates(targetRange)) {
//         const charts = activeSheet.charts;
//         charts.load("items/name,items/title");
//         await context.sync();

//         const targetChart = charts.items.find(
//           (c: any) => c.name.toLowerCase() === targetRange.toLowerCase()
//         );
//         if (targetChart) {
//           const chartImage = targetChart.getImage();
//           targetChart.title.load("text");
//           await context.sync();

//           isChart = true;
//           sheetName = activeSheet.name;
//           rangeAddress = targetRange;
//           dataSnapshot = `data:image/png;base64,${chartImage.value}`;
//           return {
//             isChart,
//             sheetName,
//             rangeAddress,
//             dataSnapshot,
//             chartTitle: targetChart.title.text || targetChart.name,
//           };
//         }
//       }

//       const range = activeSheet.getRange(targetRange);
//       range.load("address");
//       await context.sync();

//       const rangeImage = range.getImage();
//       await context.sync();

//       sheetName = activeSheet.name;
//       rangeAddress = targetRange;
//       dataSnapshot = `data:image/png;base64,${rangeImage.value}`;
//       return { isChart, sheetName, rangeAddress, dataSnapshot };
//     }

//     try {
//       const activeChart = context.workbook.getSelectedChart();
//       activeChart.load(["name", "worksheet", "title"]);
//       activeChart.title.load("text");
//       await context.sync();

//       const chartImage = activeChart.getImage();
//       await context.sync();

//       isChart = true;
//       sheetName = activeChart.worksheet.name;
//       rangeAddress = activeChart.name;
//       dataSnapshot = `data:image/png;base64,${chartImage.value}`;
//       return {
//         isChart,
//         sheetName,
//         rangeAddress,
//         dataSnapshot,
//         chartTitle: activeChart.title.text || activeChart.name,
//       };
//     } catch (chartErr) {
//       console.log("[DEBUG] Active chart selection failed, falling back to Range check.");
//     }

//     try {
//       const activeRange = context.workbook.getSelectedRange();
//       activeRange.load(["address", "worksheet"]);
//       await context.sync();

//       const rangeImage = activeRange.getImage();
//       await context.sync();

//       sheetName = activeRange.worksheet.name;
//       const fullAddress = activeRange.address;
//       rangeAddress = fullAddress.includes("!") ? fullAddress.split("!")[1] : fullAddress;
//       rangeAddress = rangeAddress.replace(/['"]/g, "");

//       dataSnapshot = `data:image/png;base64,${rangeImage.value}`;
//       return { isChart, sheetName, rangeAddress, dataSnapshot };
//     } catch (rangeErr) {
//       try {
//         const activeSheet = context.workbook.worksheets.getActiveWorksheet();
//         activeSheet.load("name");
//         const charts = activeSheet.charts;
//         charts.load("items/name,items/title");
//         await context.sync();

//         if (charts.items.length > 0) {
//           const firstChart = charts.items[0];
//           const chartImage = firstChart.getImage();
//           firstChart.title.load("text");
//           await context.sync();

//           isChart = true;
//           sheetName = activeSheet.name;
//           rangeAddress = firstChart.name;
//           dataSnapshot = `data:image/png;base64,${chartImage.value}`;
//           return {
//             isChart,
//             sheetName,
//             rangeAddress,
//             dataSnapshot,
//             chartTitle: firstChart.title.text || firstChart.name,
//           };
//         }
//       } catch (innerErr) {}

//       throw new Error(
//         "No active range or chart detected. Please select a range or click on the chart container."
//       );
//     }
//   });
// };

// export const formatExcelRange = async (_sheetName: string, _rangeAddress: string) => {
// };

// export const clearExcelRangeFormat = async (_sheetName: string, _rangeAddress: string) => {
// };

// export const saveMetadataToCustomXml = async (
//   fileId: string,
//   linkId: string,
//   sheetName: string,
//   rangeAddress: string,
//   type: "Table" | "Chart"
// ) => {
//   const xmlString = `<LiveLink><LinkId>${linkId}</LinkId><FileId>${fileId}</FileId><SheetName>${sheetName}</SheetName><RangeAddress>${rangeAddress}</RangeAddress><Type>${type}</Type></LiveLink>`;

//   await Excel.run(async (context: any) => {
//     context.workbook.customXmlParts.add(xmlString);
//     await context.sync();
//     console.log(`[DEBUG] Successfully wrote Custom XML for Link ID: ${linkId}`);
//   });
// };

declare const Excel: any;

export interface LinkMatch {
  linkId: string;
  matchedRange: string;
}

interface ExcelRect {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
}

export const generateUUID = (): string => {
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    const v = c === "x" ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
};

export const getFileNameFromUrl = (url: string): string => {
  if (!url || url === "excel-local-livelink") return "Unsaved Workbook.xlsx";
  try {
    const decodedUrl = decodeURIComponent(url);
    const parts = decodedUrl.split(/[\\/]/);
    return parts[parts.length - 1] || "Workbook.xlsx";
  } catch (e) {
    return "Workbook.xlsx";
  }
};

const colLetterToNumber = (letter: string): number => {
  let num = 0;
  for (let i = 0; i < letter.length; i++) {
    num = num * 26 + (letter.charCodeAt(i) - 64);
  }
  return num;
};

const parseRangeAddress = (address: string): ExcelRect => {
  const cleanAddress = address.replace(/[$]/g, "");
  const parts = cleanAddress.split(":");

  const parseCell = (cell: string) => {
    const match = cell.match(/^([A-Z]+)([0-9]+)$/i);
    if (!match) return { row: 0, col: 0 };
    return {
      row: parseInt(match[2], 10),
      col: colLetterToNumber(match[1].toUpperCase()),
    };
  };

  const start = parseCell(parts[0]);
  const end = parts[1] ? parseCell(parts[1]) : start;

  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  };
};

const rangesIntersect = (rect1: ExcelRect, rect2: ExcelRect): boolean => {
  return !(
    rect1.startRow > rect2.endRow ||
    rect1.endRow < rect2.startRow ||
    rect1.startCol > rect2.endCol ||
    rect1.endCol < rect2.startCol
  );
};

const isCoordinates = (str: string): boolean => {
  return /^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$/i.test(str.replace(/[$]/g, ""));
};

export const getExistingLinkId = async (
  sheetName: string,
  rangeAddress: string
): Promise<LinkMatch | null> => {
  return await Excel.run(async (context: any) => {
    const parts = context.workbook.customXmlParts;
    parts.load("items");
    await context.sync();

    // PropertyNotLoaded se bachne ke liye baghair zaroorat 'id' ko map nahi kiya gaya
    const xmlBlobs = parts.items.map((part: any) => {
      return {
        xmlBlob: part.getXml(),
      };
    });
    await context.sync();

    const parser = new DOMParser();
    const cleanSheet = sheetName.trim().toLowerCase();
    const cleanRange = rangeAddress.trim().toLowerCase();

    for (let item of xmlBlobs) {
      try {
        const xmlText = item.xmlBlob.value;
        if (xmlText && xmlText.includes("LiveLink")) {
          const xmlDoc = parser.parseFromString(xmlText, "text/xml");
          const savedSheet = (xmlDoc.getElementsByTagName("SheetName")[0]?.textContent || "")
            .trim()
            .toLowerCase();
          const savedRange = (xmlDoc.getElementsByTagName("RangeAddress")[0]?.textContent || "")
            .trim()
            .toLowerCase();
          const savedLinkId = xmlDoc.getElementsByTagName("LinkId")[0]?.textContent || "";

          if (savedSheet === cleanSheet) {
            if (savedRange === cleanRange) {
              return {
                linkId: savedLinkId,
                matchedRange: savedRange,
              };
            }

            if (isCoordinates(savedRange) && isCoordinates(cleanRange)) {
              const savedRect = parseRangeAddress(savedRange);
              const currentRect = parseRangeAddress(cleanRange);

              if (rangesIntersect(currentRect, savedRect)) {
                return {
                  linkId: savedLinkId,
                  matchedRange: savedRange,
                };
              }
            }
          }
        }
      } catch (partErr) {
        console.error("Failed to parse single Custom XML Part safely:", partErr);
      }
    }

    for (let item of xmlBlobs) {
      try {
        const xmlText = item.xmlBlob.value;
        if (xmlText && xmlText.includes("LiveLink")) {
          const xmlDoc = parser.parseFromString(xmlText, "text/xml");
          const savedSheet = (xmlDoc.getElementsByTagName("SheetName")[0]?.textContent || "")
            .trim()
            .toLowerCase();
          const savedRange = (xmlDoc.getElementsByTagName("RangeAddress")[0]?.textContent || "")
            .trim()
            .toLowerCase();
          const savedLinkId = xmlDoc.getElementsByTagName("LinkId")[0]?.textContent || "";
          const savedType = (xmlDoc.getElementsByTagName("Type")[0]?.textContent || "").trim();

          if (savedSheet === cleanSheet && savedType === "Chart") {
            return {
              linkId: savedLinkId,
              matchedRange: savedRange,
            };
          }
        }
      } catch (e) {
        console.error("Failed to scan fallback chart links:", e);
      }
    }

    return null;
  });
};

export const getActiveSelection = async (
  targetRange?: string
): Promise<{
  isChart: boolean;
  sheetName: string;
  rangeAddress: string;
  dataSnapshot: any;
  chartTitle?: string;
}> => {
  return await Excel.run(async (context: any) => {
    let isChart = false;
    let sheetName = "";
    let rangeAddress = "";
    let dataSnapshot: any = null;

    if (targetRange) {
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load("name");
      await context.sync();

      if (!isCoordinates(targetRange)) {
        const charts = activeSheet.charts;
        charts.load("items/name,items/title");
        await context.sync();

        const targetChart = charts.items.find(
          (c: any) => c.name.toLowerCase() === targetRange.toLowerCase()
        );
        if (targetChart) {
          const chartImage = targetChart.getImage();
          targetChart.title.load("text");
          await context.sync();

          isChart = true;
          sheetName = activeSheet.name;
          rangeAddress = targetRange;
          dataSnapshot = `data:image/png;base64,${chartImage.value}`;
          return {
            isChart,
            sheetName,
            rangeAddress,
            dataSnapshot,
            chartTitle: targetChart.title.text || targetChart.name,
          };
        }
      }

      const range = activeSheet.getRange(targetRange);
      range.load("address");
      await context.sync();

      const rangeImage = range.getImage();
      await context.sync();

      sheetName = activeSheet.name;
      rangeAddress = targetRange;
      dataSnapshot = `data:image/png;base64,${rangeImage.value}`;
      return { isChart, sheetName, rangeAddress, dataSnapshot };
    }

    // --- CHART SELECTION ---
    try {
      const activeChart = context.workbook.getSelectedChart();
      activeChart.load("name");
      
      const activeWorksheet = activeChart.worksheet;
      activeWorksheet.load("name"); // Worksheet ka naam explicitly load kiya gaya hai

      const chartTitleObj = activeChart.title;
      chartTitleObj.load("text");
      
      await context.sync();

      const chartImage = activeChart.getImage();
      await context.sync();

      isChart = true;
      sheetName = activeWorksheet.name;
      rangeAddress = activeChart.name;
      dataSnapshot = `data:image/png;base64,${chartImage.value}`;
      
      let chartTitle = activeChart.name;
      try {
        if (chartTitleObj && chartTitleObj.text) {
          chartTitle = chartTitleObj.text;
        }
      } catch (titleErr) {
        // Safe fallback agar title set na ho
      }

      return {
        isChart,
        sheetName,
        rangeAddress,
        dataSnapshot,
        chartTitle,
      };
    } catch (chartErr) {
      console.log("[DEBUG] Active chart selection failed, falling back to Range check.");
    }

    // --- RANGE SELECTION ---
    try {
      const activeRange = context.workbook.getSelectedRange();
      activeRange.load("address");
      
      const activeWorksheet = activeRange.worksheet;
      activeWorksheet.load("name"); // Worksheet ka naam explicitly load kiya gaya hai
      
      await context.sync();

      const rangeImage = activeRange.getImage();
      await context.sync();

      sheetName = activeWorksheet.name;
      const fullAddress = activeRange.address;
      rangeAddress = fullAddress.includes("!") ? fullAddress.split("!")[1] : fullAddress;
      rangeAddress = rangeAddress.replace(/['"]/g, "");

      dataSnapshot = `data:image/png;base64,${rangeImage.value}`;
      return { isChart, sheetName, rangeAddress, dataSnapshot };
    } catch (rangeErr) {
      try {
        const activeSheet = context.workbook.worksheets.getActiveWorksheet();
        activeSheet.load("name");
        const charts = activeSheet.charts;
        charts.load("items/name,items/title");
        await context.sync();

        if (charts.items.length > 0) {
          const firstChart = charts.items[0];
          const chartImage = firstChart.getImage();
          firstChart.title.load("text");
          await context.sync();

          isChart = true;
          sheetName = activeSheet.name;
          rangeAddress = firstChart.name;
          dataSnapshot = `data:image/png;base64,${chartImage.value}`;
          return {
            isChart,
            sheetName,
            rangeAddress,
            dataSnapshot,
            chartTitle: firstChart.title.text || firstChart.name,
          };
        }
      } catch (innerErr) {}

      throw new Error(
        "No active range or chart detected. Please select a range or click on the chart container."
      );
    }
  });
};

export const formatExcelRange = async (_sheetName: string, _rangeAddress: string) => {
};

export const clearExcelRangeFormat = async (_sheetName: string, _rangeAddress: string) => {
};

export const saveMetadataToCustomXml = async (
  fileId: string,
  linkId: string,
  sheetName: string,
  rangeAddress: string,
  type: "Table" | "Chart"
) => {
  const xmlString = `<LiveLink><LinkId>${linkId}</LinkId><FileId>${fileId}</FileId><SheetName>${sheetName}</SheetName><RangeAddress>${rangeAddress}</RangeAddress><Type>${type}</Type></LiveLink>`;

  await Excel.run(async (context: any) => {
    context.workbook.customXmlParts.add(xmlString);
    await context.sync();
    console.log(`[DEBUG] Successfully wrote Custom XML for Link ID: ${linkId}`);
  });
};