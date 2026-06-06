// declare const Office: any;
// declare const Excel: any;
// import React, { useEffect, useState } from "react";
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
//   const [linking, setLinking] = useState<boolean>(false);

//   const [isRangeLinked, setIsRangeLinked] = useState<string | null>(null);
//   const [matchedRangeAddress, setMatchedRangeAddress] = useState<string | null>(null);

//   const [customName, setCustomName] = useState<string>("");
//   const [fetchingName, setFetchingName] = useState<boolean>(false);

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
//     try {
//       const selection = await getActiveSelection();
//       const match = await getExistingLinkId(selection.sheetName, selection.rangeAddress);

//       if (match) {
//         setFetchingName(true);
//         try {
//           const res = await getLinkDetails(match.linkId);
//           if (res.success && res.data) {
//             setIsRangeLinked(match.linkId);
//             setMatchedRangeAddress(match.matchedRange);
//             setCustomName(res.data.componentName || "");
//           } else {
//             setIsRangeLinked(null);
//             setMatchedRangeAddress(null);
//             setCustomName("");
//           }
//         } catch (dbErr) {
//           console.warn("MongoDB verification failed:", dbErr);
//           setIsRangeLinked(null);
//           setMatchedRangeAddress(null);
//           setCustomName("");
//         } finally {
//           setFetchingName(false);
//         }
//       } else {
//         setIsRangeLinked(null);
//         setMatchedRangeAddress(null);
//         setCustomName("");
//       }
//     } catch (e) {
//       setIsRangeLinked(null);
//       setMatchedRangeAddress(null);
//       setCustomName("");
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
//               componentName: customName,
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
//       setCustomName("");
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

      
//         <Box sx={{ width: "100%", display: "flex", flexDirection: "column", alignItems: "center", mb: 1 }}>
//           <TextField
//             size="small"
//             label="Custom Name"
//             placeholder="e.g. Monthly Revenue Table"
//             value={customName}
//             disabled={linking || fetchingName}
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
//                   disabled={linking || !customName.trim() || fetchingName}
//                   onClick={handleCreateLiveLink}
//                   endIcon={
//                     linking ? (
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
//                   {linking ? "Updating..." : "Update Data"}
//                 </Button>
//               </span>
//             </Tooltip>

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
//                   disabled={linking || !customName.trim() || fetchingName}
//                   onClick={handleCreateLiveLink}
//                   endIcon={
//                     linking ? (
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
//                   {linking ? "Linking..." : "Send to PowerPoint"}
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

import { registerLinkData, deleteLinkData, getLinkDetails } from "../services/api";
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

const Dashboard: React.FC<DashboardProps> = () => {
  const [statusMessage, setStatusMessage] = useState<{
    text: string;
    severity: "success" | "error" | "info";
  } | null>(null);
  const [linking, setLinking] = useState<boolean>(false);

  const [isRangeLinked, setIsRangeLinked] = useState<string | null>(null);
  const [matchedRangeAddress, setMatchedRangeAddress] = useState<string | null>(null);

  const [customName, setCustomName] = useState<string>("a");
  const [fetchingName, setFetchingName] = useState<boolean>(false);
  
  // State for locking selection events during input focus [1]
  const [isInputFocused, setIsInputFocused] = useState<boolean>(false);

  useEffect(() => {
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
  }, [isInputFocused]); // Dependency injected

  const handleSelectionChanged = async () => {
    // If input field is focused, completely ignore selection changes to prevent focus deselect crashes [1]
    if (isInputFocused) {
      console.log("[DEBUG] Selection change ignored because input is focused.");
      return;
    }

    try {
      const selection = await getActiveSelection();
      const match = await getExistingLinkId(selection.sheetName, selection.rangeAddress);

      if (match) {
        setIsRangeLinked(match.linkId);
        setMatchedRangeAddress(match.matchedRange);

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

        <Box sx={{ width: "100%", display: "flex", flexDirection: "column", alignItems: "center", mb: 2 }}>
          <TextField
            size="small"
            label="Custom Name"
            placeholder="e.g. Monthly Revenue Table"
            value={customName}
            disabled={linking || fetchingName}
            // Bind onFocus and onBlur to safely lock focus-loss deselects [1]
            onFocus={() => setIsInputFocused(true)}
            onBlur={() => setTimeout(() => setIsInputFocused(false), 300)}
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
                  disabled={linking || !customName.trim() || fetchingName}
                  onClick={handleCreateLiveLink}
                  endIcon={
                    linking ? (
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
                width: "70%", 
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
                  disabled={linking || !customName.trim() || fetchingName}
                  onClick={handleCreateLiveLink}
                  endIcon={
                    linking ? (
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