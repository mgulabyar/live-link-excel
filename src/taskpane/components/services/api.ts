import axios from "axios";

// const API_BASE_URL = "http://localhost:5000/api/links";
const API_BASE_URL = "https://live-link-backend.vercel.app/api/links";

const apiInstance = axios.create({
  baseURL: API_BASE_URL,
  headers: {
    "Content-Type": "application/json",
  },
});

export const registerLinkData = async (payload: {
  linkId: string;
  excelFileId: string;
  excelFileName: string;
  sheetName: string;
  rangeAddress: string;
  type: "Table" | "Chart";
  dataSnapshot: any;
}) => {
  const response = await apiInstance.post("/register", payload);
  return response.data;
};

export const deleteLinkData = async (linkId: string) => {
  const response = await apiInstance.delete(`/${linkId}`);
  return response.data;
};