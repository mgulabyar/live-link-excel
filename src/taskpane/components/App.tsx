declare const Office: any;

import React, { useEffect, useState } from "react";
import { ThemeProvider, createTheme, Box, CircularProgress, CssBaseline } from "@mui/material";
import Dashboard from "./main/Dashboard"; 

const theme = createTheme({
  palette: {
    primary: { main: "#0078d4" },
  },
});

const App: React.FC = () => {
  const [officeInitialized, setOfficeInitialized] = useState<boolean>(false);

  useEffect(() => {
    Office.onReady(() => {
      setOfficeInitialized(true);
    });
  }, []);

  if (!officeInitialized) {
    return (
      <Box
        sx={{ display: "flex", height: "100vh", justifyContent: "center", alignItems: "center" }}
      >
        <CircularProgress color="primary" />
      </Box>
    );
  }

  return (
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <Dashboard onLogout={() => {}} />
    </ThemeProvider>
  );
};

export default App;