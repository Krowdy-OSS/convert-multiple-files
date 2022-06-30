"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.convertToBase64 = exports.convertWordFileToHTML = exports.convertWordFiles = void 0;
var path = require("path");
var child_process = require("child_process");
var mammoth = require("mammoth-style");
var fs = require("fs");
var convertWordFiles = function (pathFile, extOutput, outputDir) { return __awaiter(void 0, void 0, void 0, function () {
    var system, extension, fileName, fullName, convertCommandLinux, convertCommandWindows;
    return __generator(this, function (_a) {
        system = process.platform;
        extension = path.extname(pathFile);
        fileName = path.basename(pathFile, extension);
        fullName = path.basename(pathFile);
        convertCommandLinux = 'timeout 6s '+ path.resolve(__dirname, 'utils', 'instdir', 'program', 'soffice.bin') + " --headless --norestore --invisible --nodefault --nofirststartwizard --nolockcheck --nologo --convert-to " + extOutput + " --outdir " + outputDir + " '" + pathFile + "'";
        convertCommandWindows = path.resolve(__dirname, 'utils', 'LibreOfficePortable', 'App', 'libreoffice', 'program', 'soffice.bin') + " --headless --norestore --invisible --nodefault --nofirststartwizard --nolockcheck --nologo --convert-to " + extOutput + " --outdir " + outputDir + " \"" + pathFile + "\"";
        if (!fullName.match(/\.(doc|docx|pdf|odt)$/)) {
            throw new Error('Invalid file format, see the documentation for more information.');
        }
        else if (!extOutput.match(/(doc|docx|pdf|odt)$/)) {
            throw new Error('Format to be converted not accepted');
        }
        try {
            if (system === 'linux') {
                child_process.execSync(convertCommandLinux).toString('utf8');
            }
            if (system === 'win32') {
                child_process.execSync(convertCommandWindows).toString('utf8');
            }
        }
        catch (e) {
            throw new Error('Error converting the file');
        }
        return [2 /*return*/, path.join(outputDir, fileName + "." + extOutput)];
    });
}); };
exports.convertWordFiles = convertWordFiles;
var convertWordFileToHTML = function (pathFile, outputDir, outputPrefix) { return __awaiter(void 0, void 0, void 0, function () {
    var contentHTML, titleTags, alignTitle, newContentHTML, e_1;
    return __generator(this, function (_a) {
        switch (_a.label) {
            case 0:
                _a.trys.push([0, 2, , 3]);
                return [4 /*yield*/, mammoth.convertToHtml({ path: pathFile })];
            case 1:
                contentHTML = (_a.sent()).value;
                if (!contentHTML)
                    return [2 /*return*/];
                if (contentHTML.search('<p>') === 0) {
                    titleTags = contentHTML.substring(0, contentHTML.indexOf('</p>') + 4);
                    alignTitle = titleTags.replace(/<p>/g, '<center>');
                    alignTitle = alignTitle.replace(/<\/p>/g, '</center>');
                    newContentHTML = contentHTML.replace(titleTags, alignTitle);
                    fs.writeFileSync(path.resolve(outputDir, outputPrefix + ".html"), newContentHTML);
                    return [2 /*return*/, {
                            output: "" + path.resolve(outputDir, outputPrefix + ".html")
                        }];
                }
                fs.writeFileSync(path.resolve(outputDir, outputPrefix + ".html"), contentHTML);
                return [2 /*return*/, {
                        output: "" + path.resolve(outputDir, outputPrefix + ".html")
                    }];
            case 2:
                e_1 = _a.sent();
                throw new Error('Error converting the file to HTML');
            case 3: return [2 /*return*/];
        }
    });
}); };
exports.convertWordFileToHTML = convertWordFileToHTML;
var convertToBase64 = function (pathFile) { return __awaiter(void 0, void 0, void 0, function () {
    var data, dataBase64;
    return __generator(this, function (_a) {
        data = fs.readFileSync(pathFile);
        dataBase64 = Buffer.from(data).toString('base64');
        return [2 /*return*/, dataBase64];
    });
}); };
exports.convertToBase64 = convertToBase64;
