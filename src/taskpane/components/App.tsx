
// declare const Office: any;

// import React, { useEffect, useState } from "react";
// import { ThemeProvider, createTheme, Box, CircularProgress, CssBaseline } from "@mui/material";
// import Login from "./pages/Login";
// import Signup from "./pages/Signup";
// import Dashboard from "./main/Dashboard"; 

// const theme = createTheme({
//   palette: {
//     primary: { main: "#0078d4" },
//   },
// });

// const App: React.FC = () => {
//   const [isLoggedIn, setIsLoggedIn] = useState<boolean | null>(null);
//   const [currentScreen, setCurrentScreen] = useState<"login" | "signup">("login");

//   useEffect(() => {
//     Office.onReady(() => {
//       const status = localStorage.getItem("cb_isLoggedIn") === "true";
//       setIsLoggedIn(status);
//     });
//   }, []);

//   if (isLoggedIn === null) {
//     return (
//       <Box
//         sx={{ display: "flex", height: "100vh", justifyContent: "center", alignItems: "center" }}
//       >
//         <CircularProgress color="primary" />
//       </Box>
//     );
//   }

//   return (
//     <ThemeProvider theme={theme}>
//       <CssBaseline />
//       {!isLoggedIn ? (
//         currentScreen === "login" ? (
//           <Login
//             onLoginSuccess={() => setIsLoggedIn(true)}
//             onNavigateToSignup={() => setCurrentScreen("signup")}
//           />
//         ) : (
//           <Signup
//             onSignupSuccess={() => setCurrentScreen("login")}
//             onNavigateToLogin={() => setCurrentScreen("login")}
//           />
//         )
//       ) : (
//         <Dashboard
//           onLogout={() => {
//             localStorage.removeItem("cb_isLoggedIn");
//             setIsLoggedIn(false);
//             setCurrentScreen("login"); 
//           }}
//         />
//       )}
//     </ThemeProvider>
//   );
// };

// export default App;

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