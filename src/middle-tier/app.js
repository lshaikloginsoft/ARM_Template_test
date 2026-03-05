require("dotenv").config();
import * as createError from "http-errors";
import * as path from "path";
import * as cookieParser from "cookie-parser";
import * as logger from "morgan";
import express from "express";

import { forwardMail } from "./msgraph-helper";
import { validateJwt } from "./ssoauth-helper";

const app = express();
const port = process.env.PORT || 3000;

app.set("port", port);

app.use(logger(process.env.NODE_ENV === "production" ? "combined" : "dev"));
app.use(cookieParser());
app.use(express.json({ limit: "1mb" }));
app.use(express.urlencoded({ limit: "1mb", extended: true }));

app.use(express.static(path.join(process.cwd(), "dist")));
console.log("Serving static from:", path.join(process.cwd(), "dist"));

const indexRouter = express.Router();
indexRouter.get("/", function (req, res) {
  res.status(404).send("Not intended for browser use");
});

// Route APIs
indexRouter.get("/runtime-config.js", function (req, res) {
  res.type("application/javascript");
  res.setHeader("Cache-Control", "no-store");
  res.send(`
    window.APP_CONFIG = {
      clientId: "${process.env.CLIENT_ID}",
      tenantId: "${process.env.TENANT_ID}",
      domain: "${process.env.APP_DOMAIN}"
    };
  `);
});

indexRouter.get("/health", function (req, res) {
  res.status(200).json({
    status: "ok",
    uptime: process.uptime(),
    timestamp: new Date().toISOString()
  });
});

indexRouter.post("/forwardMail", validateJwt, forwardMail);


app.use("/", indexRouter);

// Catch 404 and forward to error handler
app.use(function (req, res, next) {
  console.warn("404:", req.originalUrl);
  next(createError(404));
});

// error handler
app.use(function (err, req, res, next) {
  console.error("500: ", err);
  // render the error page
  res.status(err.status || 500).send({
    message: err.message,
  });
});

app.listen(port, () => console.log("Server listening on port: " + port));