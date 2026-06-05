// declare const Office: any;
// declare const Excel: any;
// import React, { useEffect, useState } from "react";
// import {
//   Box,
//   Typography,
//   Button,
//   CircularProgress,
//   Alert,
//   Dialog,
//   DialogTitle,
//   DialogContent,
//   DialogActions,
// } from "@mui/material";
// import { WarningAmber, Link as LinkIcon, Send, Sync, LinkOff } from "@mui/icons-material";

// import { registerLinkData, deleteLinkData } from "../services/api";
// import {
//   generateUUID,
//   getActiveSelection,
//   formatExcelRange,
//   clearExcelRangeFormat,
//   saveMetadataToCustomXml,
//   getExistingLinkId,
//   getFileNameFromUrl,
// } from "../utils/officeHelpers";

// interface DashboardProps {
//   onLogout: () => void;
// }

// const normalizeUrl = (url: string): string => {
//   if (!url) return "";
//   try {
//     const decodedUrl = decodeURIComponent(url);
//     return decodedUrl
//       .replace(/[\\/]+/g, "/")
//       .trim()
//       .toLowerCase();
//   } catch (e) {
//     return url
//       .replace(/[\\/]+/g, "/")
//       .trim()
//       .toLowerCase();
//   }
// };

// const Dashboard: React.FC<DashboardProps> = () => {
//   const [statusMessage, setStatusMessage] = useState<{
//     text: string;
//     severity: "success" | "error" | "info";
//   } | null>(null);
//   const [linking, setLinking] = useState<boolean>(false);
//   const [showCopyPrompt, setShowCopyPrompt] = useState<boolean>(false);

//   const [isRangeLinked, setIsRangeLinked] = useState<string | null>(null);
//   const [matchedRangeAddress, setMatchedRangeAddress] = useState<string | null>(null);

//   useEffect(() => {
//     checkForCopyDetection();

//     let eventResult: any;
//     Excel.run(async (context: any) => {
//       eventResult = context.workbook.onSelectionChanged.add(handleSelectionChanged);
//       await context.sync();
//       await handleSelectionChanged();
//     }).catch((err: any) => console.error("Event registration failed:", err));

//     return () => {
//       if (eventResult) {
//         Excel.run(async (context: any) => {
//           eventResult.remove();
//           await context.sync();
//         }).catch((err: any) => console.error("Event removal failed:", err));
//       }
//     };
//   }, []);

//   const handleSelectionChanged = async () => {
//     try {
//       const selection = await getActiveSelection();
//       const match = await getExistingLinkId(selection.sheetName, selection.rangeAddress);

//       if (match) {
//         setIsRangeLinked(match.linkId);
//         setMatchedRangeAddress(match.matchedRange);
//       } else {
//         setIsRangeLinked(null);
//         setMatchedRangeAddress(null);
//       }
//     } catch (e) {
//       setIsRangeLinked(null);
//       setMatchedRangeAddress(null);
//     }
//   };

//   const checkForCopyDetection = async () => {
//     try {
//       Office.context.document.getFilePropertiesAsync(async (fileResult: any) => {
//         const currentUrl = fileResult.value.url || "LocalWorkbook";

//         await Excel.run(async (context: any) => {
//           const parts = context.workbook.customXmlParts;
//           parts.load("items");
//           await context.sync();

//           const xmlBlobs = parts.items.map((part: any) => {
//             return {
//               xmlBlob: part.getXml(),
//             };
//           });
//           await context.sync();

//           let copyDetected = false;
//           const parser = new DOMParser();

//           for (let item of xmlBlobs) {
//             const xmlText = item.xmlBlob.value;

//             if (xmlText && xmlText.includes("LiveLink")) {
//               const xmlDoc = parser.parseFromString(xmlText, "text/xml");
//               const savedUrl = xmlDoc.getElementsByTagName("FileId")[0]?.textContent || "";

//               if (
//                 savedUrl &&
//                 savedUrl !== "excel-local-livelink" &&
//                 normalizeUrl(savedUrl) !== normalizeUrl(currentUrl)
//               ) {
//                 copyDetected = true;
//               }
//             }
//           }

//           if (copyDetected) {
//             setShowCopyPrompt(true);
//           }
//         });
//       });
//     } catch (err) {
//       console.error("Error scanning copy detection properties:", err);
//     }
//   };

//   const handleKeepLinks = async () => {
//     try {
//       setLinking(true);
//       Office.context.document.getFilePropertiesAsync(async (fileResult: any) => {
//         const currentUrl = fileResult.value.url || "LocalWorkbook";

//         await Excel.run(async (context: any) => {
//           const parts = context.workbook.customXmlParts;
//           parts.load("items");
//           await context.sync();

//           const xmlBlobs = parts.items.map((part: any) => ({
//             part: part,
//             xmlBlob: part.getXml(),
//           }));
//           await context.sync();

//           const parser = new DOMParser();
//           for (let item of xmlBlobs) {
//             const xmlText = item.xmlBlob.value;
//             if (xmlText && xmlText.includes("LiveLink")) {
//               const xmlDoc = parser.parseFromString(xmlText, "text/xml");
//               const linkId = xmlDoc.getElementsByTagName("LinkId")[0]?.textContent || "";
//               const sheetName = xmlDoc.getElementsByTagName("SheetName")[0]?.textContent || "";
//               const rangeAddress =
//                 xmlDoc.getElementsByTagName("RangeAddress")[0]?.textContent || "";
//               const type = xmlDoc.getElementsByTagName("Type")[0]?.textContent || "Table";

//               const newXmlString = `<LiveLink><LinkId>${linkId}</LinkId><FileId>${currentUrl}</FileId><SheetName>${sheetName}</SheetName><RangeAddress>${rangeAddress}</RangeAddress><Type>${type}</Type></LiveLink>`;
//               item.part.delete();
//               context.workbook.customXmlParts.add(newXmlString);
//             }
//           }

//           await context.sync();
//           setShowCopyPrompt(false);
//           setLinking(false);
//           setStatusMessage({
//             text: "Excel links successfully mapped to this new copy.",
//             severity: "success",
//           });
//         });
//       });
//     } catch (err) {
//       console.error("Error keeping links:", err);
//       setLinking(false);
//     }
//   };

//   const handleCreateLiveLink = async () => {
//     setLinking(true);
//     setStatusMessage(null);

//     try {
//       const selection = await getActiveSelection(matchedRangeAddress || undefined);
//       const type = selection.isChart ? "Chart" : "Table";

//       let linkId = isRangeLinked;
//       let isNewLink = false;

//       if (!linkId) {
//         linkId = generateUUID();
//         isNewLink = true;
//       }

//       Office.context.document.getFilePropertiesAsync(async (fileResult: any) => {
//         const currentUrl = fileResult.value.url || "excel-local-livelink";
//         const fileName = getFileNameFromUrl(currentUrl);

//         if (isNewLink) {
//           await saveMetadataToCustomXml(
//             currentUrl,
//             linkId!,
//             selection.sheetName,
//             selection.rangeAddress,
//             type
//           );

//           if (!selection.isChart) {
//             await formatExcelRange(selection.sheetName, selection.rangeAddress);
//           }
//         }

//         setTimeout(async () => {
//           try {
//             await registerLinkData({
//               linkId: linkId!,
//               excelFileId: currentUrl,
//               excelFileName: fileName,
//               sheetName: selection.sheetName,
//               rangeAddress: selection.rangeAddress,
//               type: type,
//               dataSnapshot: selection.dataSnapshot,
//             });

//             setIsRangeLinked(linkId);
//             setStatusMessage({
//               text: isNewLink ? "Linked successfully!" : "Link updated successfully!",
//               severity: "success",
//             });
//           } catch (apiErr: any) {
//             console.error("API Link Error:", apiErr);
//             setStatusMessage({
//               text:
//                 apiErr.response?.data?.message || apiErr.message || "Database Connection Error.",
//               severity: "error",
//             });
//           } finally {
//             setLinking(false);
//           }
//         }, 1500);
//       });
//     } catch (err: any) {
//       console.error("Linking failed:", err);
//       setStatusMessage({
//         text:
//           err.response?.data?.message || err.message || "An error occurred while linking range.",
//         severity: "error",
//       });
//       setLinking(false);
//     }
//   };

//   const handleUnlinkRange = async () => {
//     if (!isRangeLinked) return;
//     setLinking(true);
//     setStatusMessage(null);

//     try {
//       const targetRangeToUnlink = matchedRangeAddress || (await getActiveSelection()).rangeAddress;
//       const selection = await getActiveSelection(targetRangeToUnlink);

//       await Excel.run(async (context: any) => {
//         const parts = context.workbook.customXmlParts;
//         parts.load("items");
//         await context.sync();

//         const xmlBlobs = parts.items.map((part: any) => ({
//           part: part,
//           xmlBlob: part.getXml(),
//         }));
//         await context.sync();

//         const parser = new DOMParser();
//         const cleanSheet = selection.sheetName.trim().toLowerCase();
//         const cleanRange = selection.rangeAddress.trim().toLowerCase();

//         for (let item of xmlBlobs) {
//           const xmlText = item.xmlBlob.value;
//           if (xmlText && xmlText.includes("LiveLink")) {
//             const xmlDoc = parser.parseFromString(xmlText, "text/xml");
//             const savedSheet = (xmlDoc.getElementsByTagName("SheetName")[0]?.textContent || "")
//               .trim()
//               .toLowerCase();
//             const savedRange = (xmlDoc.getElementsByTagName("RangeAddress")[0]?.textContent || "")
//               .trim()
//               .toLowerCase();

//             if (savedSheet === cleanSheet && savedRange === cleanRange) {
//               if (!selection.isChart) {
//                 await clearExcelRangeFormat(selection.sheetName, selection.rangeAddress);
//               }

//               await deleteLinkData(isRangeLinked!);

//               item.part.delete();
//               await context.sync();
//               break;
//             }
//           }
//         }
//       });

//       setIsRangeLinked(null);
//       setMatchedRangeAddress(null);
//       setStatusMessage({
//         text: "Link deleted successfully!.",
//         severity: "success",
//       });
//     } catch (err: any) {
//       console.error("Unlinking failed:", err);
//       setStatusMessage({
//         text: err.message || "Failed to unlink range.",
//         severity: "error",
//       });
//     } finally {
//       setLinking(false);
//     }
//   };

//   const handleResetTemplate = async () => {
//     try {
//       setLinking(true);
//       await Excel.run(async (context: any) => {
//         const parts = context.workbook.customXmlParts;
//         parts.load("items");
//         await context.sync();

//         const xmlBlobs = parts.items.map((part: any) => {
//           return {
//             part: part,
//             xmlBlob: part.getXml(),
//           };
//         });
//         await context.sync();

//         const parser = new DOMParser();
//         for (let item of xmlBlobs) {
//           try {
//             const xmlText = item.xmlBlob.value;
//             if (xmlText && xmlText.includes("LiveLink")) {
//               const xmlDoc = parser.parseFromString(xmlText, "text/xml");
//               const sheetName = xmlDoc.getElementsByTagName("SheetName")[0]?.textContent || "";
//               const rangeAddress =
//                 xmlDoc.getElementsByTagName("RangeAddress")[0]?.textContent || "";
//               const type = xmlDoc.getElementsByTagName("Type")[0]?.textContent || "Table";
//               const linkId = xmlDoc.getElementsByTagName("LinkId")[0]?.textContent || "";

//               if (type === "Table" && sheetName && rangeAddress) {
//                 await clearExcelRangeFormat(sheetName, rangeAddress);
//               }

//               try {
//                 await deleteLinkData(linkId);
//               } catch (dbErr) {
//                 console.warn(
//                   "MongoDB unlink failed during reset, proceeding with local clear:",
//                   dbErr
//                 );
//               }

//               item.part.delete();
//             }
//           } catch (innerErr) {
//             console.error("Failed to reset single XML Part:", innerErr);
//           }
//         }

//         await context.sync();
//         setShowCopyPrompt(false);
//         setIsRangeLinked(null);
//         setMatchedRangeAddress(null);
//         setLinking(false);
//         setStatusMessage({
//           text: "Template reset completed.",
//           severity: "success",
//         });
//       });
//     } catch (err) {
//       console.error("Error resetting template:", err);
//       setLinking(false);
//     }
//   };

//   return (
//     <Box
//       sx={{
//         height: "100vh",
//         display: "flex",
//         flexDirection: "column",
//         bgcolor: "#FFFFFF",
//         fontFamily: "Segoe UI, Arial, sans-serif",
//         overflow: "hidden",
//       }}
//     >
//       <Dialog 
//         open={showCopyPrompt} 
//         disableEscapeKeyDown
//         maxWidth="xs"
//         fullWidth
//         PaperProps={{
//           sx: {
//             borderRadius: "8px",
//             p: 1.5,
//             boxShadow: "0px 8px 30px rgba(0, 0, 0, 0.15)",
//           }
//         }}
//       >
//         <DialogTitle
//           sx={{
//             display: "flex",
//             alignItems: "center",
//             gap: 1.5,
//             color: "#d32f2f",
//             fontWeight: 800,
//             fontFamily: "Segoe UI, Arial",
//             p: 1.5,
//             fontSize: "18px",
//           }}
//         >
//           <WarningAmber sx={{ fontSize: "24px" }} /> File Copy Detected
//         </DialogTitle>
//         <DialogContent sx={{ p: 1.5, overflowY: "hidden" }}>
//           <Typography
//             variant="body2"
//             sx={{
//               lineHeight: 1.6,
//               color: "#323130",
//               fontFamily: "Segoe UI, Arial",
//               fontSize: "13.5px",
//             }}
//           >
//             This workbook appears to be duplicated from an existing live-linked report template.
//             Please choose how to proceed:
//           </Typography>
//           <Box sx={{ mt: 2.5, display: "flex", flexDirection: "column", gap: 1.5 }}>
//             <Typography
//               sx={{
//                 fontFamily: "Segoe UI, Arial",
//                 fontSize: "12.5px",
//                 color: "#605E5C",
//                 lineHeight: 1.5,
//               }}
//             >
//               • <strong>Keep Links (Option A):</strong> Map connections to this new file copy [1].
//             </Typography>
//             <Typography
//               sx={{
//                 fontFamily: "Segoe UI, Arial",
//                 fontSize: "12.5px",
//                 color: "#605E5C",
//                 lineHeight: 1.5,
//               }}
//             >
//               • <strong>Reset Template (Option B):</strong> Wipe all metadata to connect this copy
//               to a new PowerPoint deck [1].
//             </Typography>
//           </Box>
//         </DialogContent>
//         <DialogActions sx={{ p: 1.5, justifyContent: "space-between", gap: 1 }}>
//           <Button
//             variant="outlined"
//             size="small"
//             onClick={handleKeepLinks}
//             disabled={linking}
//             sx={{
//               textTransform: "none",
//               fontWeight: 700,
//               fontFamily: "Segoe UI, Arial",
//               height: "36px",
//               px: 2,
//               fontSize: "13px",
//             }}
//           >
//             Keep Links
//           </Button>
//           <Button
//             variant="contained"
//             color="error"
//             size="small"
//             onClick={handleResetTemplate}
//             disabled={linking}
//             sx={{
//               textTransform: "none",
//               fontWeight: 700,
//               fontFamily: "Segoe UI, Arial",
//               height: "36px",
//               px: 2,
//               fontSize: "13px",
//               boxShadow: "none",
//             }}
//           >
//             Reset All
//           </Button>
//         </DialogActions>
//       </Dialog>

//       <Box
//         sx={{
//           position: "sticky",
//           top: 0,
//           zIndex: 100,
//           p: 2,
//           bgcolor: "#0078d4",
//           color: "#FFFFFF",
//           display: "flex",
//           alignItems: "center",
//           justifyContent: "center",
//         }}
//       >
//         <Typography
//           sx={{
//             fontWeight: 700,
//             fontSize: "16px",
//             letterSpacing: "0.5px",
//             fontFamily: "Segoe UI, Arial",
//             textAlign: "center",
//           }}
//         >
//           EXCEL TO POWERPOINT
//         </Typography>
//       </Box>

//       <Box
//         sx={{
//           p: 2,
//           flex: 1,
//           overflowY: "auto",
//           display: "flex",
//           flexDirection: "column",
//           gap: 2,
//           maxHeight: "95vh",
//           "&::-webkit-scrollbar": { display: "none" },
//           msOverflowStyle: "none",
//           scrollbarWidth: "none",
//         }}
//       >
//         <Box sx={{ display: "flex", flexDirection: "column", alignItems: "center", mb: 3.5, mt: 1 }}>
//           <LinkIcon sx={{ color: "#0078d4", fontSize: 32, mb: 0.8 }} />
//           <Typography
//             sx={{
//               fontWeight: 800,
//               fontSize: "17px",
//               color: "#323130",
//               fontFamily: "Segoe UI, Arial",
//               letterSpacing: "0.3px"
//             }}
//           >
//             Live Link
//           </Typography>
//         </Box>

//         <Typography
//           sx={{
//             fontSize: "13px",
//             color: "#605E5C",
//             lineHeight: 1.5,
//             mb: 3.5,
//             fontFamily: "Segoe UI, Arial",
//             textAlign: "center",
//             px: 1,
//           }}
//         >
//           Select any data range or chart in your sheet to create a live link. The linked object
//           can be refreshed directly in PowerPoint.
//         </Typography>

//         {isRangeLinked ? (
//           <Box sx={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 1.5, width: "100%" }}>
//             <Button
//               variant="contained"
//               disabled={linking}
//               onClick={handleCreateLiveLink}
//               endIcon={linking ? <CircularProgress size={18} color="inherit" /> : <Sync sx={{ fontSize: 18 }} />} // Unified Sync icon [1]
//               sx={{
//                 width: "70%",
//                 height: "44px",
//                 bgcolor: "#0078d4",
//                 fontWeight: 700,
//                 textTransform: "none",
//                 fontSize: "14px",
//                 boxShadow: "none",
//                 fontFamily: "Segoe UI, Arial",
//                 "&:hover": { bgcolor: "#005a9e", boxShadow: "none" },
//               }}
//             >
//               {linking ? "Updating..." : "Update Data"}
//             </Button>
//             <Button
//               variant="outlined"
//               color="error"
//               disabled={linking}
//               onClick={handleUnlinkRange}
//               endIcon={<LinkOff sx={{ fontSize: 18 }} />} 
//               sx={{
//                 width: "70%", 
//                 height: "44px",
//                 fontWeight: 700,
//                 textTransform: "none",
//                 fontSize: "14px",
//                 fontFamily: "Segoe UI, Arial",
//               }}
//             >
//               Unlink Range
//             </Button>
//           </Box>
//         ) : (
//           <Box sx={{ display: "flex", flexDirection: "column", alignItems: "center", width: "100%" }}>
//             <Button
//               variant="contained"
//               disabled={linking}
//               onClick={handleCreateLiveLink}
//               endIcon={linking ? <CircularProgress size={18} color="inherit" /> : <Send sx={{ fontSize: 18 }} />} // Customized Send icon [1]
//               sx={{
//                 width: "70%", 
//                 height: "44px",
//                 bgcolor: "#0078d4",
//                 fontWeight: 700,
//                 textTransform: "none",
//                 fontSize: "14px",
//                 boxShadow: "none",
//                 fontFamily: "Segoe UI, Arial",
//                 "&:hover": { bgcolor: "#005a9e", boxShadow: "none" },
//               }}
//             >
//               {linking ? "Linking..." : "Send to PowerPoint"}
//             </Button>
//           </Box>
//         )}

//         {statusMessage && (
//           <Alert
//             severity={statusMessage.severity}
//             sx={{ mt: 2.5, fontSize: "13px", fontFamily: "Segoe UI, Arial", textAlign: "left" }}
//           >
//             {statusMessage.text}
//           </Alert>
//         )}
//       </Box>

//       <Box sx={{ p: 1.5, textAlign: "center", borderTop: "1px solid #EDEBE9" }}>
//         <Typography
//           sx={{ fontSize: "10px", color: "#A19F9D", fontWeight: 600, fontFamily: "Segoe UI, Arial" }}
//         >
//           Live Linker v1.0.0
//         </Typography>
//       </Box>
//     </Box>
//   );
// };

// export default Dashboard;


declare const Office: any;
declare const Excel: any;
import React, { useEffect, useState } from "react";
import {
  Box,
  Typography,
  Button,
  CircularProgress,
  Alert,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  Snackbar,
  TextField,
  Tooltip,
} from "@mui/material";
import { WarningAmber, Link as LinkIcon, Send, Sync, LinkOff } from "@mui/icons-material";

import { registerLinkData, deleteLinkData, getLinkDetails } from "../services/api"; // getLinkDetails imported
import {
  generateUUID,
  getActiveSelection,
  formatExcelRange,
  clearExcelRangeFormat,
  saveMetadataToCustomXml,
  getExistingLinkId,
  getFileNameFromUrl,
} from "../utils/officeHelpers";

interface DashboardProps {
  onLogout: () => void;
}

const normalizeUrl = (url: string): string => {
  if (!url) return "";
  try {
    const decodedUrl = decodeURIComponent(url);
    return decodedUrl
      .replace(/[\\/]+/g, "/")
      .trim()
      .toLowerCase();
  } catch (e) {
    return url
      .replace(/[\\/]+/g, "/")
      .trim()
      .toLowerCase();
  }
};

const Dashboard: React.FC<DashboardProps> = () => {
  const [statusMessage, setStatusMessage] = useState<{
    text: string;
    severity: "success" | "error" | "info";
  } | null>(null);
  const [linking, setLinking] = useState<boolean>(false);
  const [showCopyPrompt, setShowCopyPrompt] = useState<boolean>(false);

  const [isRangeLinked, setIsRangeLinked] = useState<string | null>(null);
  const [matchedRangeAddress, setMatchedRangeAddress] = useState<string | null>(null);
  
  // Custom naming states
  const [customName, setCustomName] = useState<string>("");
  const [fetchingName, setFetchingName] = useState<boolean>(false);

  useEffect(() => {
    checkForCopyDetection();

    let eventResult: any;
    Excel.run(async (context: any) => {
      eventResult = context.workbook.onSelectionChanged.add(handleSelectionChanged);
      await context.sync();
      await handleSelectionChanged();
    }).catch((err: any) => console.error("Event registration failed:", err));

    return () => {
      if (eventResult) {
        Excel.run(async (context: any) => {
          eventResult.remove();
          await context.sync();
        }).catch((err: any) => console.error("Event removal failed:", err));
      }
    };
  }, []);

  const handleSelectionChanged = async () => {
    try {
      const selection = await getActiveSelection();
      const match = await getExistingLinkId(selection.sheetName, selection.rangeAddress);

      if (match) {
        setIsRangeLinked(match.linkId);
        setMatchedRangeAddress(match.matchedRange);
        
        // Database se is linked range ka custom name fetch karein [1]
        setFetchingName(true);
        try {
          const res = await getLinkDetails(match.linkId);
          if (res.success && res.data) {
            setCustomName(res.data.componentName || "");
          }
        } catch (dbErr) {
          console.error("Failed to fetch link custom name:", dbErr);
        } finally {
          setFetchingName(false);
        }
      } else {
        setIsRangeLinked(null);
        setMatchedRangeAddress(null);
        setCustomName("");
      }
    } catch (e) {
      setIsRangeLinked(null);
      setMatchedRangeAddress(null);
      setCustomName("");
    }
  };

  const checkForCopyDetection = async () => {
    try {
      Office.context.document.getFilePropertiesAsync(async (fileResult: any) => {
        const currentUrl = fileResult.value.url || "LocalWorkbook";

        await Excel.run(async (context: any) => {
          const parts = context.workbook.customXmlParts;
          parts.load("items");
          await context.sync();

          const xmlBlobs = parts.items.map((part: any) => {
            return {
              xmlBlob: part.getXml(),
            };
          });
          await context.sync();

          let copyDetected = false;
          const parser = new DOMParser();

          for (let item of xmlBlobs) {
            const xmlText = item.xmlBlob.value;

            if (xmlText && xmlText.includes("LiveLink")) {
              const xmlDoc = parser.parseFromString(xmlText, "text/xml");
              const savedUrl = xmlDoc.getElementsByTagName("FileId")[0]?.textContent || "";

              if (
                savedUrl &&
                savedUrl !== "excel-local-livelink" &&
                normalizeUrl(savedUrl) !== normalizeUrl(currentUrl)
              ) {
                copyDetected = true;
              }
            }
          }

          if (copyDetected) {
            setShowCopyPrompt(true);
          }
        });
      });
    } catch (err) {
      console.error("Error scanning copy detection properties:", err);
    }
  };

  const handleKeepLinks = async () => {
    try {
      setLinking(true);
      Office.context.document.getFilePropertiesAsync(async (fileResult: any) => {
        const currentUrl = fileResult.value.url || "LocalWorkbook";

        await Excel.run(async (context: any) => {
          const parts = context.workbook.customXmlParts;
          parts.load("items");
          await context.sync();

          const xmlBlobs = parts.items.map((part: any) => ({
            part: part,
            xmlBlob: part.getXml(),
          }));
          await context.sync();

          const parser = new DOMParser();
          for (let item of xmlBlobs) {
            const xmlText = item.xmlBlob.value;
            if (xmlText && xmlText.includes("LiveLink")) {
              const xmlDoc = parser.parseFromString(xmlText, "text/xml");
              const linkId = xmlDoc.getElementsByTagName("LinkId")[0]?.textContent || "";
              const sheetName = xmlDoc.getElementsByTagName("SheetName")[0]?.textContent || "";
              const rangeAddress =
                xmlDoc.getElementsByTagName("RangeAddress")[0]?.textContent || "";
              const type = xmlDoc.getElementsByTagName("Type")[0]?.textContent || "Table";

              const newXmlString = `<LiveLink><LinkId>${linkId}</LinkId><FileId>${currentUrl}</FileId><SheetName>${sheetName}</SheetName><RangeAddress>${rangeAddress}</RangeAddress><Type>${type}</Type></LiveLink>`;
              item.part.delete();
              context.workbook.customXmlParts.add(newXmlString);
            }
          }

          await context.sync();
          setShowCopyPrompt(false);
          setLinking(false);
          setStatusMessage({
            text: "Excel links successfully mapped to this new copy.",
            severity: "success",
          });
        });
      });
    } catch (err) {
      console.error("Error keeping links:", err);
      setLinking(false);
    }
  };

  const handleCreateLiveLink = async () => {
    setLinking(true);
    setStatusMessage(null);

    try {
      const selection = await getActiveSelection(matchedRangeAddress || undefined);
      const type = selection.isChart ? "Chart" : "Table";

      let linkId = isRangeLinked;
      let isNewLink = false;

      if (!linkId) {
        linkId = generateUUID();
        isNewLink = true;
      }

      Office.context.document.getFilePropertiesAsync(async (fileResult: any) => {
        const currentUrl = fileResult.value.url || "excel-local-livelink";
        const fileName = getFileNameFromUrl(currentUrl);

        if (isNewLink) {
          await saveMetadataToCustomXml(
            currentUrl,
            linkId!,
            selection.sheetName,
            selection.rangeAddress,
            type
          );

          if (!selection.isChart) {
            await formatExcelRange(selection.sheetName, selection.rangeAddress);
          }
        }

        setTimeout(async () => {
          try {
            await registerLinkData({
              linkId: linkId!,
              componentName: customName, 
              excelFileId: currentUrl,
              excelFileName: fileName,
              sheetName: selection.sheetName,
              rangeAddress: selection.rangeAddress,
              type: type,
              dataSnapshot: selection.dataSnapshot,
            });

            setIsRangeLinked(linkId);
            setStatusMessage({
              text: isNewLink ? "Linked successfully!" : "Link updated successfully!",
              severity: "success",
            });
          } catch (apiErr: any) {
            console.error("API Link Error:", apiErr);
            setStatusMessage({
              text:
                apiErr.response?.data?.message || apiErr.message || "Database Connection Error.",
              severity: "error",
            });
          } finally {
            setLinking(false);
          }
        }, 1500);
      });
    } catch (err: any) {
      console.error("Linking failed:", err);
      setStatusMessage({
        text:
          err.response?.data?.message || err.message || "An error occurred while linking range.",
        severity: "error",
      });
      setLinking(false);
    }
  };

  const handleUnlinkRange = async () => {
    if (!isRangeLinked) return;
    setLinking(true);
    setStatusMessage(null);

    try {
      const targetRangeToUnlink = matchedRangeAddress || (await getActiveSelection()).rangeAddress;
      const selection = await getActiveSelection(targetRangeToUnlink);

      await Excel.run(async (context: any) => {
        const parts = context.workbook.customXmlParts;
        parts.load("items");
        await context.sync();

        const xmlBlobs = parts.items.map((part: any) => ({
          part: part,
          xmlBlob: part.getXml(),
        }));
        await context.sync();

        const parser = new DOMParser();
        const cleanSheet = selection.sheetName.trim().toLowerCase();
        const cleanRange = selection.rangeAddress.trim().toLowerCase();

        for (let item of xmlBlobs) {
          const xmlText = item.xmlBlob.value;
          if (xmlText && xmlText.includes("LiveLink")) {
            const xmlDoc = parser.parseFromString(xmlText, "text/xml");
            const savedSheet = (xmlDoc.getElementsByTagName("SheetName")[0]?.textContent || "")
              .trim()
              .toLowerCase();
            const savedRange = (xmlDoc.getElementsByTagName("RangeAddress")[0]?.textContent || "")
              .trim()
              .toLowerCase();

            if (savedSheet === cleanSheet && savedRange === cleanRange) {
              if (!selection.isChart) {
                await clearExcelRangeFormat(selection.sheetName, selection.rangeAddress);
              }

              await deleteLinkData(isRangeLinked!);

              item.part.delete();
              await context.sync();
              break;
            }
          }
        }
      });

      // FIXED: Instant local state clearing to dynamically revert UI back to Send to PowerPoint mode [1]
      setIsRangeLinked(null);
      setMatchedRangeAddress(null);
      setCustomName(""); 
      setStatusMessage({
        text: "Link deleted successfully!.",
        severity: "success",
      });
    } catch (err: any) {
      console.error("Unlinking failed:", err);
      setStatusMessage({
        text: err.message || "Failed to unlink range.",
        severity: "error",
      });
    } finally {
      setLinking(false);
    }
  };

  const handleResetTemplate = async () => {
    try {
      setLinking(true);
      await Excel.run(async (context: any) => {
        const parts = context.workbook.customXmlParts;
        parts.load("items");
        await context.sync();

        const xmlBlobs = parts.items.map((part: any) => {
          return {
            part: part,
            xmlBlob: part.getXml(),
          };
        });
        await context.sync();

        const parser = new DOMParser();
        for (let item of xmlBlobs) {
          try {
            const xmlText = item.xmlBlob.value;
            if (xmlText && xmlText.includes("LiveLink")) {
              const xmlDoc = parser.parseFromString(xmlText, "text/xml");
              const sheetName = xmlDoc.getElementsByTagName("SheetName")[0]?.textContent || "";
              const rangeAddress =
                xmlDoc.getElementsByTagName("RangeAddress")[0]?.textContent || "";
              const type = xmlDoc.getElementsByTagName("Type")[0]?.textContent || "Table";
              const linkId = xmlDoc.getElementsByTagName("LinkId")[0]?.textContent || "";

              if (type === "Table" && sheetName && rangeAddress) {
                await clearExcelRangeFormat(sheetName, rangeAddress);
              }

              try {
                await deleteLinkData(linkId);
              } catch (dbErr) {
                console.warn(
                  "MongoDB unlink failed during reset, proceeding with local clear:",
                  dbErr
                );
              }

              item.part.delete();
            }
          } catch (innerErr) {
            console.error("Failed to reset single XML Part:", innerErr);
          }
        }

        await context.sync();
        setShowCopyPrompt(false);
        setIsRangeLinked(null);
        setMatchedRangeAddress(null);
        setCustomName("");
        setLinking(false);
        setStatusMessage({
          text: "Template reset completed.",
          severity: "success",
        });
      });
    } catch (err) {
      console.error("Error resetting template:", err);
      setLinking(false);
    }
  };

  return (
    <Box
      sx={{
        height: "100vh",
        display: "flex",
        flexDirection: "column",
        bgcolor: "#FFFFFF",
        fontFamily: "Segoe UI, Arial, sans-serif",
        overflow: "hidden",
      }}
    >
      <Dialog 
        open={showCopyPrompt} 
        disableEscapeKeyDown
        maxWidth="xs"
        fullWidth
        PaperProps={{
          sx: {
            borderRadius: "8px",
            p: 1.5,
            boxShadow: "0px 8px 30px rgba(0, 0, 0, 0.15)",
          }
        }}
      >
        <DialogTitle
          sx={{
            display: "flex",
            alignItems: "center",
            gap: 1.5,
            color: "#d32f2f",
            fontWeight: 800,
            fontFamily: "Segoe UI, Arial",
            p: 1.5,
            fontSize: "18px",
          }}
        >
          <WarningAmber sx={{ fontSize: "24px" }} /> File Copy Detected
        </DialogTitle>
        <DialogContent sx={{ p: 1.5, overflowY: "hidden" }}>
          <Typography
            variant="body2"
            sx={{
              lineHeight: 1.6,
              color: "#323130",
              fontFamily: "Segoe UI, Arial",
              fontSize: "13.5px",
            }}
          >
            This workbook appears to be duplicated from an existing live-linked report template.
            Please choose how to proceed:
          </Typography>
          <Box sx={{ mt: 2.5, display: "flex", flexDirection: "column", gap: 1.5 }}>
            <Typography
              sx={{
                fontFamily: "Segoe UI, Arial",
                fontSize: "12.5px",
                color: "#605E5C",
                lineHeight: 1.5,
              }}
            >
              • <strong>Keep Links (Option A):</strong> Map connections to this new file copy [1].
            </Typography>
            <Typography
              sx={{
                fontFamily: "Segoe UI, Arial",
                fontSize: "12.5px",
                color: "#605E5C",
                lineHeight: 1.5,
              }}
            >
              • <strong>Reset Template (Option B):</strong> Wipe all metadata to connect this copy
              to a new PowerPoint deck [1].
            </Typography>
          </Box>
        </DialogContent>
        <DialogActions sx={{ p: 1.5, justifyContent: "space-between", gap: 1 }}>
          <Button
            variant="outlined"
            size="small"
            onClick={handleKeepLinks}
            disabled={linking}
            sx={{
              textTransform: "none",
              fontWeight: 700,
              fontFamily: "Segoe UI, Arial",
              height: "36px",
              px: 2,
              fontSize: "13px",
            }}
          >
            Keep Links
          </Button>
          <Button
            variant="contained"
            color="error"
            size="small"
            onClick={handleResetTemplate}
            disabled={linking}
            sx={{
              textTransform: "none",
              fontWeight: 700,
              fontFamily: "Segoe UI, Arial",
              height: "36px",
              px: 2,
              fontSize: "13px",
              boxShadow: "none",
            }}
          >
            Reset All
          </Button>
        </DialogActions>
      </Dialog>

      <Box
        sx={{
          position: "sticky",
          top: 0,
          zIndex: 100,
          p: 2,
          bgcolor: "#0078d4",
          color: "#FFFFFF",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
        }}
      >
        <Typography
          sx={{
            fontWeight: 700,
            fontSize: "16px",
            letterSpacing: "0.5px",
            fontFamily: "Segoe UI, Arial",
            textAlign: "center",
          }}
        >
          EXCEL TO POWERPOINT
        </Typography>
      </Box>

      <Box
        sx={{
          p: 2,
          flex: 1,
          overflowY: "auto",
          display: "flex",
          flexDirection: "column",
          gap: 2,
          maxHeight: "95vh",
          "&::-webkit-scrollbar": { display: "none" },
          msOverflowStyle: "none",
          scrollbarWidth: "none",
        }}
      >
        <Box sx={{ display: "flex", flexDirection: "column", alignItems: "center", mb: 2, mt: 1 }}>
          <LinkIcon sx={{ color: "#0078d4", fontSize: 32, mb: 0.8 }} />
          <Typography
            sx={{
              fontWeight: 800,
              fontSize: "17px",
              color: "#323130",
              fontFamily: "Segoe UI, Arial",
              letterSpacing: "0.3px"
            }}
          >
            Live Link
          </Typography>
        </Box>

        <Typography
          sx={{
            fontSize: "13px",
            color: "#605E5C",
            lineHeight: 1.5,
            mb: 2.5,
            fontFamily: "Segoe UI, Arial",
            textAlign: "center",
            px: 1,
          }}
        >
          Select any data range or chart in your sheet to create a live link. The linked object
          can be refreshed directly in PowerPoint.
        </Typography>

        {/* CUSTOM NAME INPUT FIELD AREA [1] */}
        <Box sx={{ width: "100%", display: "flex", flexDirection: "column", alignItems: "center", mb: 2 }}>
          <TextField
            fullWidth
            size="small"
            label="Custom Name"
            placeholder="e.g. Monthly Revenue Table"
            value={customName}
            disabled={linking || fetchingName}
            onChange={(e) => setCustomName(e.target.value)}
            sx={{
              width: "90%",
              "& .MuiOutlinedInput-root": { height: "42px", fontSize: "13px", fontFamily: "Segoe UI, Arial" },
              "& .MuiInputLabel-root": { fontSize: "13px", fontFamily: "Segoe UI, Arial" }
            }}
            InputProps={{
              endAdornment: fetchingName && <CircularProgress size={16} color="inherit" />
            }}
          />
        </Box>

        {/* DYNAMIC BUTTONS PANEL */}
        {isRangeLinked ? (
          <Box sx={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 1.5, width: "100%" }}>
            {/* Dynamic Update button with tooltipped validation */}
            <Tooltip 
              title={!customName.trim() ? "Please enter a custom name to update." : "Update data on slide"} 
              arrow 
              placement="top"
            >
              <span>
                <Button
                  variant="contained"
                  disabled={linking || !customName.trim() || fetchingName}
                  onClick={handleCreateLiveLink}
                  endIcon={linking ? <CircularProgress size={18} color="inherit" /> : <Sync sx={{ fontSize: 18 }} />}
                  sx={{
                    width: "90%",
                    height: "44px",
                    bgcolor: "#0078d4",
                    fontWeight: 700,
                    textTransform: "none",
                    fontSize: "14px",
                    boxShadow: "none",
                    fontFamily: "Segoe UI, Arial",
                    "&:hover": { bgcolor: "#005a9e", boxShadow: "none" },
                  }}
                >
                  {linking ? "Updating..." : "Update Data"}
                </Button>
              </span>
            </Tooltip>

            <Button
              variant="outlined"
              color="error"
              disabled={linking}
              onClick={handleUnlinkRange}
              endIcon={<LinkOff sx={{ fontSize: 18 }} />} 
              sx={{
                width: "90%", 
                height: "44px",
                fontWeight: 700,
                textTransform: "none",
                fontSize: "14px",
                fontFamily: "Segoe UI, Arial",
              }}
            >
              Unlink Range
            </Button>
          </Box>
        ) : (
          <Box sx={{ display: "flex", flexDirection: "column", alignItems: "center", width: "100%" }}>
            {/* Dynamic Send button with tooltipped validation [1] */}
            <Tooltip 
              title={!customName.trim() ? "Please enter a custom name first." : "Send snapshot to slide"} 
              arrow 
              placement="top"
            >
              <span>
                <Button
                  variant="contained"
                  disabled={linking || !customName.trim() || fetchingName}
                  onClick={handleCreateLiveLink}
                  endIcon={linking ? <CircularProgress size={18} color="inherit" /> : <Send sx={{ fontSize: 18 }} />}
                  sx={{
                    width: "90%", 
                    height: "44px",
                    bgcolor: "#0078d4",
                    fontWeight: 700,
                    textTransform: "none",
                    fontSize: "14px",
                    boxShadow: "none",
                    fontFamily: "Segoe UI, Arial",
                    "&:hover": { bgcolor: "#005a9e", boxShadow: "none" },
                  }}
                >
                  {linking ? "Linking..." : "Send to PowerPoint"}
                </Button>
              </span>
            </Tooltip>
          </Box>
        )}
      </Box>

      <Snackbar
        open={statusMessage !== null}
        autoHideDuration={2000} 
        onClose={() => setStatusMessage(null)}
        anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
      >
        {statusMessage ? (
          <Alert
            onClose={() => setStatusMessage(null)}
            severity={statusMessage.severity}
            sx={{ width: "100%", fontSize: "13px", fontFamily: "Segoe UI, Arial" }}
          >
            {statusMessage.text}
          </Alert>
        ) : undefined}
      </Snackbar>

      <Box sx={{ p: 1.5, textAlign: "center", borderTop: "1px solid #EDEBE9" }}>
        <Typography
          sx={{ fontSize: "10px", color: "#A19F9D", fontWeight: 600, fontFamily: "Segoe UI, Arial" }}
        >
          Live Linker v1.0.0
        </Typography>
      </Box>
    </Box>
  );
};

export default Dashboard;