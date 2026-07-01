import React from "react";
import ReactDOM from "react-dom/client";
import "./storage.js"; // installs a localStorage-backed window.storage before the app reads it
import "./index.css";
import App from "./App.jsx";

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
