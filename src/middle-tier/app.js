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

// view engine setup
app.set("views", path.join(__dirname, "views"));
app.set("view engine", "pug");

app.use(logger("dev"));
app.use(cookieParser());
app.use(express.json({ limit: "100mb" }));
app.use(express.urlencoded({ limit: "100mb", extended: true }));

app.use(express.static(path.join(process.cwd(), "dist")));
console.log("Serving static from:", path.join(process.cwd(), "dist"));

const indexRouter = express.Router();
indexRouter.get("/", function (req, res) {
  res.sendFile(path.join(process.cwd(), "dist", "taskpane.html"));
});

// Route APIs
indexRouter.post("/forwardMail", validateJwt, forwardMail);

app.use("/", indexRouter);

// Catch 404 and forward to error handler
app.use(function (req, res, next) {
  console.log("error 404");
  next(createError(404));
});

// error handler
app.use(function (err, req, res, next) {
  // set locals, only providing error in development
  console.log("error 500");
  res.locals.message = err.message;
  res.locals.error = req.app.get("env") === "development" ? err : {};

  // render the error page
  res.status(err.status || 500).send({
    message: err.message,
  });
});

app.listen(port, () => console.log("Server listening on port: " + port));