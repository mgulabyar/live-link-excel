import React, { useState, useEffect } from "react";
import {
  Box,
  Typography,
  TextField,
  Button,
  InputAdornment,
  IconButton,
  Link,
  Divider,
  Snackbar,
  Alert,
  CircularProgress,
} from "@mui/material";
import { Email, Lock, Visibility, VisibilityOff, TableChart } from "@mui/icons-material";

interface SignupProps {
  onSignupSuccess: () => void;
  onNavigateToLogin: () => void;
}

const Signup: React.FC<SignupProps> = ({ onSignupSuccess, onNavigateToLogin }) => {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [showPass, setShowPass] = useState(false);
  const [showConfirmPass, setShowConfirmPass] = useState(false);
  const [authLoading, setAuthLoading] = useState(false);

  const [toast, setToast] = useState({
    open: false,
    message: "",
    severity: "error" as "error" | "success" | "info",
  });

  useEffect(() => {
    if (localStorage.getItem("cb_redirectToSignup") === "true") {
      setToast({
        open: true,
        message: "Create an account to continue.",
        severity: "info",
      });
      localStorage.removeItem("cb_redirectToSignup");
    }
  }, []);

  const handleAuth = () => {
    localStorage.setItem("user_creds", JSON.stringify({ email, password }));
    setToast({
      open: true,
      message: "Account created! Redirecting...",
      severity: "success",
    });
    localStorage.setItem("cb_justSignedUp", "true"); 
    
    setTimeout(() => {
      onSignupSuccess();
    }, 1500);
  };

  const handleAuthWithLoading = () => {
    if (!email || !password || !confirmPassword) {
      setToast({ open: true, message: "Please fill in all fields.", severity: "error" });
      return;
    }

    if (password !== confirmPassword) {
      setToast({ open: true, message: "Passwords do not match.", severity: "error" });
      return;
    }

    setAuthLoading(true);

    setTimeout(() => {
      setAuthLoading(false);
      handleAuth();
    }, 1500);
  };

  const inputSx = {
    "& .MuiOutlinedInput-root": {
      height: "44px",
      fontSize: "14px",
      borderRadius: "4px",
      bgcolor: "#ffffff",
      fontFamily: "Arial, sans-serif",
      "& fieldset": {
        borderColor: "#D1D1D1",
      },
      "&:hover fieldset": {
        borderColor: "#0078d4",
      },
      "&.Mui-focused fieldset": {
        borderColor: "#0078d4",
        borderWidth: "2px",
      },
    },
    "& .MuiInputLabel-root.Mui-focused": {
      color: "#0078d4",
      fontFamily: "Arial, sans-serif",
    },
    "& .MuiInputLabel-root": { fontSize: "14px", fontFamily: "Arial, sans-serif" },
    "& input": { fontFamily: "Arial, sans-serif" },
    "& input::-ms-reveal, & input::-ms-clear": { display: "none" },
  };

  return (
    <Box
      sx={{
        height: "100vh",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        bgcolor: "#ffffff",
        overflow: "hidden",
        p: 3,
        fontFamily: "Arial, sans-serif",
      }}
    >
      <Box
        sx={{
          width: "100%",
          maxWidth: "310px",
          textAlign: "center",
          maxHeight: "95vh", 
          overflowY: "auto", 
          "&::-webkit-scrollbar": { display: "none" }, 
          msOverflowStyle: "none", 
          scrollbarWidth: "none", 
        }}
      >
        <Box sx={{ mb: 3 }}>
          <TableChart sx={{ fontSize: 50, color: "#0078d4" }} />
          <Typography variant="h6" sx={{ fontWeight: 700, color: "#323130", fontFamily: "Arial", lineHeight: 1, fontSize: "24px" }}>
            Excel To PowerPoint
          </Typography>
          <Typography sx={{ fontSize: "10px", color: "#605E5C", letterSpacing: "1px", fontWeight: 700, mt: 0.5, fontFamily: "Arial" }}>
            LIVE LINKING ENGINE
          </Typography>
        </Box>

        <Box sx={{ display: "flex", flexDirection: "column", gap: 2.2 }}>
          <TextField
            fullWidth
            size="small"
            placeholder="Enter Email"
            sx={inputSx}
            type="email"
            value={email}
            disabled={authLoading}
            onChange={(e) => setEmail(e.target.value)}
            slotProps={{
              input: {
                startAdornment: (
                  <InputAdornment position="start">
                    <Email sx={{ fontSize: 19, color: "#605E5C" }} />
                  </InputAdornment>
                ),
              },
            }}
          />

          <TextField
            fullWidth
            size="small"
            placeholder="Password"
            type={showPass ? "text" : "password"}
            sx={inputSx}
            value={password}
            disabled={authLoading}
            onChange={(e) => setPassword(e.target.value)}
            slotProps={{
              input: {
                startAdornment: (
                  <InputAdornment position="start">
                    <Lock sx={{ fontSize: 19, color: "#605E5C" }} />
                  </InputAdornment>
                ),
                endAdornment: (
                  <InputAdornment position="end">
                    <IconButton onClick={() => setShowPass(!showPass)} edge="end" size="small" sx={{ p: "3px" }} disabled={authLoading}>
                      {showPass ? <VisibilityOff sx={{ fontSize: 19 }} /> : <Visibility sx={{ fontSize: 19 }} />}
                    </IconButton>
                  </InputAdornment>
                ),
              },
            }}
          />

          <TextField
            fullWidth
            size="small"
            placeholder="Confirm Password"
            type={showConfirmPass ? "text" : "password"}
            sx={inputSx}
            value={confirmPassword}
            disabled={authLoading}
            onChange={(e) => setConfirmPassword(e.target.value)}
            slotProps={{
              input: {
                startAdornment: (
                  <InputAdornment position="start">
                    <Lock sx={{ fontSize: 19, color: "#605E5C" }} />
                  </InputAdornment>
                ),
                endAdornment: (
                  <InputAdornment position="end">
                    <IconButton onClick={() => setShowConfirmPass(!showConfirmPass)} edge="end" size="small" sx={{ p: "3px" }} disabled={authLoading}>
                      {showConfirmPass ? <VisibilityOff sx={{ fontSize: 19 }} /> : <Visibility sx={{ fontSize: 19 }} />}
                    </IconButton>
                  </InputAdornment>
                ),
              },
            }}
          />

          <Button
            variant="contained"
            fullWidth
            onClick={handleAuthWithLoading}
            disabled={authLoading}
            sx={{
              height: "40px",
              fontSize: "14px",
              fontWeight: 700,
              textTransform: "none",
              borderRadius: "4px",
              boxShadow: "none",
              bgcolor: "#0078d4",
              fontFamily: "Arial",
              "&:hover": { bgcolor: "#005a9e", boxShadow: "none" },
              mt: 1,
            }}
          >
            {authLoading ? (
              <CircularProgress size={20} thickness={5} sx={{ color: "#ffffff", "& .MuiCircularProgress-circle": { strokeDasharray: "4px, 4px" } }} />
            ) : (
              "Sign Up"
            )}
          </Button>
        </Box>

        <Divider sx={{ my: 2.5 }}>
          <Typography sx={{ color: "#A19F9D", fontSize: "12px", fontWeight: 700, px: 2, fontFamily: "Arial" }}>OR</Typography>
        </Divider>

        <Box sx={{ mb: 2, pb: 4, display: "flex", alignItems: "center", justifyContent: "center", gap: 0.5 }}>
          <Typography variant="caption" color="text.secondary" sx={{ fontSize: "13px", fontFamily: "Arial" }}>
            Already have an account?
          </Typography>
          <Link
            component="button"
            variant="caption"
            onClick={onNavigateToLogin}
            sx={{ fontWeight: 700, color: "#0078d4", textDecoration: "none", fontSize: "13px", fontFamily: "Arial" }}
          >
            Log in
          </Link>
        </Box>
      </Box>

      <Snackbar open={toast.open} autoHideDuration={3000} onClose={() => setToast({ ...toast, open: false })} anchorOrigin={{ vertical: "bottom", horizontal: "center" }}>
        <Alert onClose={() => setToast({ ...toast, open: false })} severity={toast.severity} sx={{ width: "100%", fontSize: "13px", fontFamily: "Arial" }}>
          {toast.message}
        </Alert>
      </Snackbar>
    </Box>
  );
};

export default Signup;