// declare const Office: any;
// declare const Excel: any;
// import React, { useEffect, useState, useRef } from "react";
// import {
//   Box,
//   Typography,
//   Button,
//   CircularProgress,
//   Alert,
//   Snackbar,
//   TextField,
//   Tooltip,
// } from "@mui/material";
// import { Link as LinkIcon, Send, Sync, LinkOff } from "@mui/icons-material";

// import { registerLinkData, deleteLinkData, getLinkDetails } from "../services/api";
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

// const Dashboard: React.FC<DashboardProps> = () => {
//   const [statusMessage, setStatusMessage] = useState<{
//     text: string;
//     severity: "success" | "error" | "info";
//   } | null>(null);

//   // Independent loading states
//   const [isLinking, setIsLinking] = useState<boolean>(false);
//   const [isUpdating, setIsUpdating] = useState<boolean>(false);
//   const [isUnlinking, setIsUnlinking] = useState<boolean>(false);

//   const [isRangeLinked, setIsRangeLinked] = useState<string | null>(null);
//   const [matchedRangeAddress, setMatchedRangeAddress] = useState<string | null>(null);

//   const [customName, setCustomName] = useState<string>("");
//   const [fetchingName, setFetchingName] = useState<boolean>(false);
  
//   // Track if a Chart is currently active to dynamically disable inputs [1]
//   const [isChartSelected, setIsChartSelected] = useState<boolean>(false);

//   // Zero-latency selection cache ref [1]
//   const [lastSelection, setLastSelection] = useState<any>(null);
//   const isMouseInPaneRef = useRef<boolean>(false);

//   useEffect(() => {
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
//     // Zero-latency check: Ignore Excel selection change if mouse is in pane [1]
//     if (isMouseInPaneRef.current) {
//       console.log("[DEBUG] Selection change ignored because mouse is active inside the taskpane.");
//       return;
//     }

//     try {
//       const selection = await getActiveSelection();
//       setLastSelection(selection); // Cache the selection state [1]
//       setIsChartSelected(selection.isChart); // Check if chart [1]

//       const match = await getExistingLinkId(selection.sheetName, selection.rangeAddress);

//       if (match) {
//         setFetchingName(true);
//         try {
//           const res = await getLinkDetails(match.linkId);
//           if (res.success && res.data) {
//             setIsRangeLinked(match.linkId);
//             setMatchedRangeAddress(match.matchedRange);
//             // Auto-load custom name or chart title [1]
//             setCustomName(res.data.componentName || selection.chartTitle || "");
//           } else {
//             setIsRangeLinked(null);
//             setMatchedRangeAddress(null);
//             setCustomName(selection.chartTitle || "");
//           }
//         } catch (dbErr) {
//           console.warn("MongoDB verification failed:", dbErr);
//           setIsRangeLinked(null);
//           setMatchedRangeAddress(null);
//           setCustomName(selection.chartTitle || "");
//         } finally {
//           setFetchingName(false);
//         }
//       } else {
//         setIsRangeLinked(null);
//         setMatchedRangeAddress(null);
//         setCustomName(selection.chartTitle || ""); // Map active Excel Chart Title to input [1]
//       }
//     } catch (e) {
//       console.log("[DEBUG] Selection temporarily lost, retaining previous active selection.");
//     }
//   };

//   const handleCreateLiveLink = async () => {
//     try {
//       // Use cached selection to bypass deselect focus-loss [1]
//       const selection = matchedRangeAddress 
//         ? await getActiveSelection(matchedRangeAddress) 
//         : (lastSelection && lastSelection.isChart) ? lastSelection : await getActiveSelection();

//       const type = selection.isChart ? "Chart" : "Table";
//       let linkId = isRangeLinked;
//       let isNewLink = false;

//       if (!linkId) {
//         linkId = generateUUID();
//         isNewLink = true;
//       }

//       if (isNewLink) {
//         setIsLinking(true);
//       } else {
//         setIsUpdating(true);
//       }
//       setStatusMessage(null);

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
//             // If it is a chart, prioritize its own Excel Chart Title over the manual customName! [1]
//             const finalComponentName = selection.isChart ? (selection.chartTitle || customName) : customName;

//             await registerLinkData({
//               linkId: linkId!,
//               componentName: finalComponentName,
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
//             setIsLinking(false);
//             setIsUpdating(false);
//           }
//         }, 1500);
//       });
//     } catch (err: any) {
//       console.error("Linking failed:", err);
//       setStatusMessage({
//         text: err.message || "An error occurred while linking range.",
//         severity: "error",
//       });
//       setIsLinking(false);
//       setIsUpdating(false);
//     }
//   };

//   const handleUnlinkRange = async () => {
//     if (!isRangeLinked) return;
//     setIsUnlinking(true);
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
//       setCustomName("");
//       setStatusMessage({
//         text: "Link deleted successfully!",
//         severity: "success",
//       });
//     } catch (err: any) {
//       console.error("Unlinking failed:", err);
//       setStatusMessage({
//         text: err.message || "Failed to unlink range.",
//         severity: "error",
//       });
//     } finally {
//       setIsUnlinking(false);
//     }
//   };

//   return (
//     <Box
//       onMouseEnter={() => { isMouseInPaneRef.current = true; }} // Active Lock [1]
//       onMouseLeave={() => { isMouseInPaneRef.current = false; }} // Active Unlock [1]
//       sx={{
//         height: "100vh",
//         display: "flex",
//         flexDirection: "column",
//         bgcolor: "#FFFFFF",
//         fontFamily: "Segoe UI, Arial, sans-serif",
//         overflow: "hidden",
//       }}
//     >
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
//         <Box
//           sx={{ display: "flex", flexDirection: "column", alignItems: "center", mb: 3.5, mt: 1 }}
//         >
//           <LinkIcon sx={{ color: "#0078d4", fontSize: 32, mb: 0.8 }} />
//           <Typography
//             sx={{
//               fontWeight: 800,
//               fontSize: "17px",
//               color: "#323130",
//               fontFamily: "Segoe UI, Arial",
//               letterSpacing: "0.3px",
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
//             mb: 2.5,
//             fontFamily: "Segoe UI, Arial",
//             textAlign: "center",
//             px: 1,
//           }}
//         >
//           Select any data range or chart in your sheet to create a live link. The linked object can
//           be refreshed directly in PowerPoint.
//         </Typography>

//         {/* Input box margin tightened by 1px (mb: 1) for professional compact visual alignment */}
//         <Box sx={{ width: "100%", display: "flex", flexDirection: "column", alignItems: "center", mb: 1 }}>
//           <TextField
//             size="small"
//             label={isChartSelected ? "Chart Title (From Excel)" : "Custom Name"} // Dynamically changes labels
//             placeholder={isChartSelected ? "Define title on Excel Chart" : "e.g. Monthly Revenue Table"}
//             value={customName}
//             // Disabled if it is a Chart (Forces user to set name in Excel title, preventing deselect) [1]
//             disabled={isChartSelected || isLinking || isUpdating || isUnlinking || fetchingName}
//             onChange={(e) => setCustomName(e.target.value)}
//             sx={{
//               width: "70%", 
//               "& .MuiOutlinedInput-root": {
//                 height: "44px", 
//                 fontSize: "13px",
//                 fontFamily: "Segoe UI, Arial",
//               },
//               "& .MuiInputLabel-root": { fontSize: "13px", fontFamily: "Segoe UI, Arial" },
//             }}
//             InputProps={{
//               endAdornment: fetchingName && <CircularProgress size={16} color="inherit" />,
//             }}
//           />
//         </Box>

//         {isRangeLinked ? (
//           <Box
//             sx={{
//               display: "flex",
//               flexDirection: "column",
//               alignItems: "center",
//               gap: 1.5,
//               width: "100%",
//             }}
//           >
//             <Tooltip
//               title={
//                 !customName.trim()
//                   ? "Please enter a custom name to update."
//                   : "Update data on slide"
//               }
//               arrow
//               placement="top"
//             >
//               <span style={{ display: "block", width: "70%" }}>
//                 <Button
//                   variant="contained"
//                   disabled={isUpdating || isUnlinking || !customName.trim() || fetchingName}
//                   onClick={handleCreateLiveLink}
//                   endIcon={
//                     isUpdating ? (
//                       <CircularProgress size={18} color="inherit" />
//                     ) : (
//                       <Sync sx={{ fontSize: 18 }} />
//                     )
//                   }
//                   sx={{
//                     width: "100%", 
//                     height: "44px",
//                     bgcolor: "#0078d4",
//                     fontWeight: 700,
//                     textTransform: "none",
//                     fontSize: "14px",
//                     boxShadow: "none",
//                     fontFamily: "Segoe UI, Arial",
//                     "&:hover": { bgcolor: "#005a9e", boxShadow: "none" },
//                   }}
//                 >
//                   {isUpdating ? "Updating..." : "Update Data"}
//                 </Button>
//               </span>
//             </Tooltip>

//             <Button
//               variant="outlined"
//               color="error"
//               disabled={isUpdating || isUnlinking}
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
//               {isUnlinking ? "Unlinking..." : "Unlink Range"}
//             </Button>
//           </Box>
//         ) : (
//           <Box
//             sx={{ display: "flex", flexDirection: "column", alignItems: "center", width: "100%" }}
//           >
//             <Tooltip
//               title={
//                 !customName.trim() ? "Please enter a custom name first." : "Send snapshot to slide"
//               }
//               arrow
//               placement="top"
//             >
//               <span style={{ display: "block", width: "70%" }}>
//                 <Button
//                   variant="contained"
//                   disabled={isLinking || !customName.trim() || fetchingName}
//                   onClick={handleCreateLiveLink}
//                   endIcon={
//                     isLinking ? (
//                       <CircularProgress size={18} color="inherit" />
//                     ) : (
//                       <Send sx={{ fontSize: 18 }} />
//                     )
//                   }
//                   sx={{
//                     width: "100%", 
//                     height: "44px",
//                     bgcolor: "#0078d4",
//                     fontWeight: 700,
//                     textTransform: "none",
//                     fontSize: "14px",
//                     boxShadow: "none",
//                     fontFamily: "Segoe UI, Arial",
//                     "&:hover": { bgcolor: "#005a9e", boxShadow: "none" },
//                   }}
//                 >
//                   {isLinking ? "Linking..." : "Send to PowerPoint"}
//                 </Button>
//               </span>
//             </Tooltip>
//           </Box>
//         )}
//       </Box>

//       <Snackbar
//         open={statusMessage !== null}
//         autoHideDuration={2000}
//         onClose={() => setStatusMessage(null)}
//         anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
//       >
//         {statusMessage ? (
//           <Alert
//             onClose={() => setStatusMessage(null)}
//             severity={statusMessage.severity}
//             sx={{ width: "100%", fontSize: "13px", fontFamily: "Segoe UI, Arial" }}
//           >
//             {statusMessage.text}
//           </Alert>
//         ) : undefined}
//       </Snackbar>

//       <Box sx={{ p: 1.5, textAlign: "center", borderTop: "1px solid #EDEBE9" }}>
//         <Typography
//           sx={{
//             fontSize: "10px",
//             color: "#A19F9D",
//             fontWeight: 600,
//             fontFamily: "Segoe UI, Arial",
//           }}
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
import React, { useEffect, useRef, useState } from "react";
import {
  Box,
  Typography,
  Button,
  CircularProgress,
  Alert,
  Snackbar,
  TextField,
  Tooltip,
} from "@mui/material";
import { Link as LinkIcon, Send, Sync, LinkOff } from "@mui/icons-material";

import { registerLinkData, deleteLinkData, getLinkDetails } from "../services/api";
import {
  generateUUID,
  getActiveSelection,
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

  const [isLinking, setIsLinking] = useState<boolean>(false);
  const [isUpdating, setIsUpdating] = useState<boolean>(false);
  const [isUnlinking, setIsUnlinking] = useState<boolean>(false);

  const [isRangeLinked, setIsRangeLinked] = useState<string | null>(null);
  const [matchedRangeAddress, setMatchedRangeAddress] = useState<string | null>(null);

  const [customName, setCustomName] = useState<string>("");
  const [fetchingName, setFetchingName] = useState<boolean>(false);
  const [isChartSelected, setIsChartSelected] = useState<boolean>(false);

  const [lastSelection, setLastSelection] = useState<any>(null);
  const currentUrlRef = useRef<string>("excel-local-livelink");

  useEffect(() => {
    let eventResult: any;
    
    Office.context.document.getFilePropertiesAsync((fileResult: any) => {
      currentUrlRef.current = fileResult?.value?.url || "excel-local-livelink";
    });

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
      setLastSelection(selection); 
      setIsChartSelected(selection.isChart); 

      const match = await getExistingLinkId(selection.sheetName, selection.rangeAddress);

      if (match) {
        setIsRangeLinked(match.linkId);
        setMatchedRangeAddress(match.matchedRange);

        setFetchingName(true);
        try {
          const res = await getLinkDetails(match.linkId);
          if (res.success && res.data) {
            const currentFileName = getFileNameFromUrl(currentUrlRef.current);

            if (res.data.excelFileName !== currentFileName) {
              await registerLinkData({
                linkId: match.linkId,
                componentName: res.data.componentName || "",
                excelFileId: currentUrlRef.current,
                excelFileName: currentFileName,
                sheetName: selection.sheetName,
                rangeAddress: selection.rangeAddress,
                type: selection.isChart ? "Chart" : "Table",
                dataSnapshot: res.data.dataSnapshot
              });
            }

            setCustomName(res.data.componentName || selection.chartTitle || "");
          }
        } catch (dbErr) {
          console.warn("MongoDB verification failed, fallback to local XML match:", dbErr);
          setCustomName(selection.chartTitle || "");
        } finally {
          setFetchingName(false);
        }
      } else {
        setIsRangeLinked(null);
        setMatchedRangeAddress(null);
        setCustomName("");
      }
    } catch (e) {
      console.log("[DEBUG] Selection lost, retaining previous active selection.");
    }
  };

  const handleCreateLiveLink = async () => {
    try {
      const selection = matchedRangeAddress 
        ? await getActiveSelection(matchedRangeAddress) 
        : (lastSelection && lastSelection.isChart) ? lastSelection : await getActiveSelection();

      const type = selection.isChart ? "Chart" : "Table";
      let linkId = isRangeLinked;
      let isNewLink = false;

      if (!linkId) {
        linkId = generateUUID();
        isNewLink = true;
      }

      if (isNewLink) {
        setIsLinking(true);
      } else {
        setIsUpdating(true);
      }
      setStatusMessage(null);

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
        }

        setTimeout(async () => {
          try {
            const finalComponentName = selection.isChart ? (selection.chartTitle || customName) : customName;

            await registerLinkData({
              linkId: linkId!,
              componentName: finalComponentName,
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
            setIsLinking(false);
            setIsUpdating(false);
          }
        }, 1500);
      });
    } catch (err: any) {
      console.error("Linking failed:", err);
      setStatusMessage({
        text: err.message || "An error occurred while linking range.",
        severity: "error",
      });
      setIsLinking(false);
      setIsUpdating(false);
    }
  };

  const handleUnlinkRange = async () => {
    if (!isRangeLinked) return;
    setIsUnlinking(true);
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

      setIsRangeLinked(null);
      setMatchedRangeAddress(null);
      setCustomName("");
      setStatusMessage({
        text: "Link deleted successfully!",
        severity: "success",
      });
    } catch (err: any) {
      console.error("Unlinking failed:", err);
      setStatusMessage({
        text: err.message || "Failed to unlink range.",
        severity: "error",
      });
    } finally {
      setIsUnlinking(false);
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
        <Box
          sx={{ display: "flex", flexDirection: "column", alignItems: "center", mb: 3.5, mt: 1 }}
        >
          <LinkIcon sx={{ color: "#0078d4", fontSize: 32, mb: 0.8 }} />
          <Typography
            sx={{
              fontWeight: 800,
              fontSize: "17px",
              color: "#323130",
              fontFamily: "Segoe UI, Arial",
              letterSpacing: "0.3px",
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
          Select any data range or chart in your sheet to create a live link. The linked object can
          be refreshed directly in PowerPoint.
        </Typography>

        <Box sx={{ width: "100%", display: "flex", flexDirection: "column", alignItems: "center", mb: 1 }}>
          <TextField
            size="small"
            label={isChartSelected ? "Chart Title (From Excel)" : "Custom Name"} 
            placeholder={isChartSelected ? "Define title on Excel Chart" : "e.g. Monthly Revenue Table"}
            value={customName}
            disabled={isChartSelected || isLinking || isUpdating || isUnlinking || fetchingName}
            onChange={(e) => setCustomName(e.target.value)}
            sx={{
              width: "70%", 
              "& .MuiOutlinedInput-root": {
                height: "44px", 
                fontSize: "13px",
                fontFamily: "Segoe UI, Arial",
              },
              "& .MuiInputLabel-root": { fontSize: "13px", fontFamily: "Segoe UI, Arial" },
            }}
            InputProps={{
              endAdornment: fetchingName && <CircularProgress size={16} color="inherit" />,
            }}
          />
        </Box>

        {isRangeLinked ? (
          <Box
            sx={{
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              gap: 1.5,
              width: "100%",
            }}
          >
            <Tooltip
              title={
                !customName.trim()
                  ? "Please enter a custom name to update."
                  : "Update data on slide"
              }
              arrow
              placement="top"
            >
              <span style={{ display: "block", width: "70%" }}>
                <Button
                  variant="contained"
                  disabled={isUpdating || isUnlinking || !customName.trim() || fetchingName}
                  onClick={handleCreateLiveLink}
                  endIcon={
                    isUpdating ? (
                      <CircularProgress size={18} color="inherit" />
                    ) : (
                      <Sync sx={{ fontSize: 18 }} />
                    )
                  }
                  sx={{
                    width: "100%", 
                    height: "44px",
                    bgcolor: "#0078d4",
                    fontWeight: 700,
                    textTransform: "none",
                    fontSize: "14px",
                    boxShadow: "none",
                    fontFamily: "Segoe UI, Arial",
                    "&:hover": { bgcolor: "#005a9e", boxShadow: "none" }
                  }}
                >
                  {isUpdating ? "Updating..." : "Update Data"}
                </Button>
              </span>
            </Tooltip>

            <Button
              variant="outlined"
              color="error"
              disabled={isUpdating || isUnlinking}
              onClick={handleUnlinkRange}
              endIcon={<LinkOff sx={{ fontSize: 18 }} />}
              sx={{
                width: "70%", 
                height: "44px",
                fontWeight: 700,
                textTransform: "none",
                fontSize: "14px",
                fontFamily: "Segoe UI, Arial",
              }}
            >
              {isUnlinking ? "Unlinking..." : "Unlink Range"}
            </Button>
          </Box>
        ) : (
          <Box
            sx={{ display: "flex", flexDirection: "column", alignItems: "center", width: "100%" }}
          >
            <Tooltip
              title={
                !customName.trim() ? "Please enter a custom name first." : "Send snapshot to slide"
              }
              arrow
              placement="top"
            >
              <span style={{ display: "block", width: "70%" }}>
                <Button
                  variant="contained"
                  disabled={isLinking || !customName.trim() || fetchingName}
                  onClick={handleCreateLiveLink}
                  endIcon={
                    isLinking ? (
                      <CircularProgress size={18} color="inherit" />
                    ) : (
                      <Send sx={{ fontSize: 18 }} />
                    )
                  }
                  sx={{
                    width: "100%", 
                    height: "44px",
                    bgcolor: "#0078d4",
                    fontWeight: 700,
                    textTransform: "none",
                    fontSize: "14px",
                    boxShadow: "none",
                    fontFamily: "Segoe UI, Arial",
                    "&:hover": { bgcolor: "#005a9e", boxShadow: "none" }
                  }}
                >
                  {isLinking ? "Linking..." : "Send to PowerPoint"}
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
          sx={{
            fontSize: "10px",
            color: "#A19F9D",
            fontWeight: 600,
            fontFamily: "Segoe UI, Arial",
          }}
        >
          Live Linker v1.0.0
        </Typography>
      </Box>
    </Box>
  );
};

export default Dashboard;