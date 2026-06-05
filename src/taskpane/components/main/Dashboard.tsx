declare const Office: any;
declare const Excel: any;
import React, { useEffect, useState } from "react";
import {
  Box,
  Typography,
  Button,
  Paper,
  IconButton,
  CircularProgress,
  Alert,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
} from "@mui/material";
import { CloudSync, Logout, Hub, WarningAmber } from "@mui/icons-material";

import { registerLinkData, deleteLinkData } from "../services/api";
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

const Dashboard: React.FC<DashboardProps> = ({ onLogout }) => {
  const [statusMessage, setStatusMessage] = useState<{
    text: string;
    severity: "success" | "error" | "info";
  } | null>(null);
  const [linking, setLinking] = useState<boolean>(false);
  const [showCopyPrompt, setShowCopyPrompt] = useState<boolean>(false);

  const [isRangeLinked, setIsRangeLinked] = useState<string | null>(null);
  const [matchedRangeAddress, setMatchedRangeAddress] = useState<string | null>(null);

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
      } else {
        setIsRangeLinked(null);
        setMatchedRangeAddress(null);
      }
    } catch (e) {
      setIsRangeLinked(null);
      setMatchedRangeAddress(null);
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
      <Dialog open={showCopyPrompt} disableEscapeKeyDown>
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
            Reset
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
          justifyContent: "space-between",
        }}
      >
        <Box sx={{ width: 32 }} />
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
        <IconButton onClick={onLogout} size="small" sx={{ color: "#FFFFFF" }}>
          <Logout sx={{ fontSize: 18 }} />
        </IconButton>
      </Box>

      <Box
        sx={{
          px: 2.5,
          py: 1.5,
          bgcolor: "#F3F2F1",
          display: "flex",
          alignItems: "center",
          gap: 1,
          borderBottom: "1px solid #EDEBE9",
        }}
      >
        <Box sx={{ width: 8, height: 8, bgcolor: "#107C10", borderRadius: "50%" }} />
        <Typography
          sx={{
            fontSize: "12px",
            fontWeight: 700,
            color: "#323130",
            fontFamily: "Segoe UI, Arial",
          }}
        >
          SYNC ENGINE ACTIVE
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
        <Paper
          elevation={0}
          sx={{ p: 3, border: "1px solid #D1D1D1", borderRadius: "8px", textAlign: "center" }}
        >
          <Box sx={{ display: "flex", flexDirection: "column", alignItems: "center", mb: 2 }}>
            <Hub sx={{ color: "#0078d4", fontSize: 24, mb: 1 }} />
            <Typography
              sx={{
                fontWeight: 800,
                fontSize: "16px",
                color: "#323130",
                fontFamily: "Segoe UI, Arial",
              }}
            >
              Live Data Linking
            </Typography>
          </Box>

          <Typography
            sx={{
              fontSize: "13px",
              color: "#605E5C",
              lineHeight: 1.5,
              mb: 3.5,
              fontFamily: "Segoe UI, Arial",
              px: 1,
            }}
          >
            Select any data range or chart in your sheet to create a live link. The linked object
            can be refreshed directly in PowerPoint.
          </Typography>

          {isRangeLinked ? (
            <Box sx={{ display: "flex", flexDirection: "column", gap: 1 }}>
              <Button
                variant="contained"
                fullWidth
                disabled={linking}
                onClick={handleCreateLiveLink}
                startIcon={linking ? <CircularProgress size={18} color="inherit" /> : <CloudSync />}
                sx={{
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
                {linking ? "Updating Link..." : "Update PowerPoint Link"}
              </Button>
              <Button
                variant="outlined"
                color="error"
                fullWidth
                disabled={linking}
                onClick={handleUnlinkRange}
                sx={{
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
            <Button
              variant="contained"
              fullWidth
              disabled={linking}
              onClick={handleCreateLiveLink}
              startIcon={linking ? <CircularProgress size={18} color="inherit" /> : <CloudSync />}
              sx={{
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
              {linking ? "Establishing Link..." : "Send to PowerPoint"}
            </Button>
          )}

          {statusMessage && (
            <Alert
              severity={statusMessage.severity}
              sx={{ mt: 2.5, fontSize: "13px", fontFamily: "Segoe UI, Arial", textAlign: "left" }}
            >
              {statusMessage.text}
            </Alert>
          )}
        </Paper>
      </Box>

      <Box sx={{ p: 2, textAlign: "center" }}>
        <Typography
          sx={{
            fontSize: "11px",
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
