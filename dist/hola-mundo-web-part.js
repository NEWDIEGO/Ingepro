define("96e22014-9346-4d3f-b740-b70bfed82558_0.0.1", ["@microsoft/sp-webpart-base"], (__WEBPACK_EXTERNAL_MODULE__3134__) => { return /******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ 3134:
/*!*********************************************!*\
  !*** external "@microsoft/sp-webpart-base" ***!
  \*********************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__3134__;

/***/ }),

/***/ 4340:
/*!**********************************************************************!*\
  !*** ./node_modules/@pnp/queryable/node_modules/tslib/tslib.es6.mjs ***!
  \**********************************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   __decorate: () => (/* binding */ __decorate)
/* harmony export */ });
/* unused harmony exports __extends, __assign, __rest, __param, __esDecorate, __runInitializers, __propKey, __setFunctionName, __metadata, __awaiter, __generator, __createBinding, __exportStar, __values, __read, __spread, __spreadArrays, __spreadArray, __await, __asyncGenerator, __asyncDelegator, __asyncValues, __makeTemplateObject, __importStar, __importDefault, __classPrivateFieldGet, __classPrivateFieldSet, __classPrivateFieldIn, __addDisposableResource, __disposeResources */
/******************************************************************************
Copyright (c) Microsoft Corporation.

Permission to use, copy, modify, and/or distribute this software for any
purpose with or without fee is hereby granted.

THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES WITH
REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF MERCHANTABILITY
AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY SPECIAL, DIRECT,
INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES WHATSOEVER RESULTING FROM
LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR
OTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR
PERFORMANCE OF THIS SOFTWARE.
***************************************************************************** */
/* global Reflect, Promise, SuppressedError, Symbol, Iterator */

var extendStatics = function(d, b) {
  extendStatics = Object.setPrototypeOf ||
      ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
      function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
  return extendStatics(d, b);
};

function __extends(d, b) {
  if (typeof b !== "function" && b !== null)
      throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
  extendStatics(d, b);
  function __() { this.constructor = d; }
  d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
}

var __assign = function() {
  __assign = Object.assign || function __assign(t) {
      for (var s, i = 1, n = arguments.length; i < n; i++) {
          s = arguments[i];
          for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p)) t[p] = s[p];
      }
      return t;
  }
  return __assign.apply(this, arguments);
}

function __rest(s, e) {
  var t = {};
  for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p) && e.indexOf(p) < 0)
      t[p] = s[p];
  if (s != null && typeof Object.getOwnPropertySymbols === "function")
      for (var i = 0, p = Object.getOwnPropertySymbols(s); i < p.length; i++) {
          if (e.indexOf(p[i]) < 0 && Object.prototype.propertyIsEnumerable.call(s, p[i]))
              t[p[i]] = s[p[i]];
      }
  return t;
}

function __decorate(decorators, target, key, desc) {
  var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
  if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
  else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
  return c > 3 && r && Object.defineProperty(target, key, r), r;
}

function __param(paramIndex, decorator) {
  return function (target, key) { decorator(target, key, paramIndex); }
}

function __esDecorate(ctor, descriptorIn, decorators, contextIn, initializers, extraInitializers) {
  function accept(f) { if (f !== void 0 && typeof f !== "function") throw new TypeError("Function expected"); return f; }
  var kind = contextIn.kind, key = kind === "getter" ? "get" : kind === "setter" ? "set" : "value";
  var target = !descriptorIn && ctor ? contextIn["static"] ? ctor : ctor.prototype : null;
  var descriptor = descriptorIn || (target ? Object.getOwnPropertyDescriptor(target, contextIn.name) : {});
  var _, done = false;
  for (var i = decorators.length - 1; i >= 0; i--) {
      var context = {};
      for (var p in contextIn) context[p] = p === "access" ? {} : contextIn[p];
      for (var p in contextIn.access) context.access[p] = contextIn.access[p];
      context.addInitializer = function (f) { if (done) throw new TypeError("Cannot add initializers after decoration has completed"); extraInitializers.push(accept(f || null)); };
      var result = (0, decorators[i])(kind === "accessor" ? { get: descriptor.get, set: descriptor.set } : descriptor[key], context);
      if (kind === "accessor") {
          if (result === void 0) continue;
          if (result === null || typeof result !== "object") throw new TypeError("Object expected");
          if (_ = accept(result.get)) descriptor.get = _;
          if (_ = accept(result.set)) descriptor.set = _;
          if (_ = accept(result.init)) initializers.unshift(_);
      }
      else if (_ = accept(result)) {
          if (kind === "field") initializers.unshift(_);
          else descriptor[key] = _;
      }
  }
  if (target) Object.defineProperty(target, contextIn.name, descriptor);
  done = true;
};

function __runInitializers(thisArg, initializers, value) {
  var useValue = arguments.length > 2;
  for (var i = 0; i < initializers.length; i++) {
      value = useValue ? initializers[i].call(thisArg, value) : initializers[i].call(thisArg);
  }
  return useValue ? value : void 0;
};

function __propKey(x) {
  return typeof x === "symbol" ? x : "".concat(x);
};

function __setFunctionName(f, name, prefix) {
  if (typeof name === "symbol") name = name.description ? "[".concat(name.description, "]") : "";
  return Object.defineProperty(f, "name", { configurable: true, value: prefix ? "".concat(prefix, " ", name) : name });
};

function __metadata(metadataKey, metadataValue) {
  if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(metadataKey, metadataValue);
}

function __awaiter(thisArg, _arguments, P, generator) {
  function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
  return new (P || (P = Promise))(function (resolve, reject) {
      function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
      function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
      function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
      step((generator = generator.apply(thisArg, _arguments || [])).next());
  });
}

function __generator(thisArg, body) {
  var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
  return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
  function verb(n) { return function (v) { return step([n, v]); }; }
  function step(op) {
      if (f) throw new TypeError("Generator is already executing.");
      while (g && (g = 0, op[0] && (_ = 0)), _) try {
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
}

var __createBinding = Object.create ? (function(o, m, k, k2) {
  if (k2 === undefined) k2 = k;
  var desc = Object.getOwnPropertyDescriptor(m, k);
  if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
  }
  Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
  if (k2 === undefined) k2 = k;
  o[k2] = m[k];
});

function __exportStar(m, o) {
  for (var p in m) if (p !== "default" && !Object.prototype.hasOwnProperty.call(o, p)) __createBinding(o, m, p);
}

function __values(o) {
  var s = typeof Symbol === "function" && Symbol.iterator, m = s && o[s], i = 0;
  if (m) return m.call(o);
  if (o && typeof o.length === "number") return {
      next: function () {
          if (o && i >= o.length) o = void 0;
          return { value: o && o[i++], done: !o };
      }
  };
  throw new TypeError(s ? "Object is not iterable." : "Symbol.iterator is not defined.");
}

function __read(o, n) {
  var m = typeof Symbol === "function" && o[Symbol.iterator];
  if (!m) return o;
  var i = m.call(o), r, ar = [], e;
  try {
      while ((n === void 0 || n-- > 0) && !(r = i.next()).done) ar.push(r.value);
  }
  catch (error) { e = { error: error }; }
  finally {
      try {
          if (r && !r.done && (m = i["return"])) m.call(i);
      }
      finally { if (e) throw e.error; }
  }
  return ar;
}

/** @deprecated */
function __spread() {
  for (var ar = [], i = 0; i < arguments.length; i++)
      ar = ar.concat(__read(arguments[i]));
  return ar;
}

/** @deprecated */
function __spreadArrays() {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) s += arguments[i].length;
  for (var r = Array(s), k = 0, i = 0; i < il; i++)
      for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++)
          r[k] = a[j];
  return r;
}

function __spreadArray(to, from, pack) {
  if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
      if (ar || !(i in from)) {
          if (!ar) ar = Array.prototype.slice.call(from, 0, i);
          ar[i] = from[i];
      }
  }
  return to.concat(ar || Array.prototype.slice.call(from));
}

function __await(v) {
  return this instanceof __await ? (this.v = v, this) : new __await(v);
}

function __asyncGenerator(thisArg, _arguments, generator) {
  if (!Symbol.asyncIterator) throw new TypeError("Symbol.asyncIterator is not defined.");
  var g = generator.apply(thisArg, _arguments || []), i, q = [];
  return i = Object.create((typeof AsyncIterator === "function" ? AsyncIterator : Object).prototype), verb("next"), verb("throw"), verb("return", awaitReturn), i[Symbol.asyncIterator] = function () { return this; }, i;
  function awaitReturn(f) { return function (v) { return Promise.resolve(v).then(f, reject); }; }
  function verb(n, f) { if (g[n]) { i[n] = function (v) { return new Promise(function (a, b) { q.push([n, v, a, b]) > 1 || resume(n, v); }); }; if (f) i[n] = f(i[n]); } }
  function resume(n, v) { try { step(g[n](v)); } catch (e) { settle(q[0][3], e); } }
  function step(r) { r.value instanceof __await ? Promise.resolve(r.value.v).then(fulfill, reject) : settle(q[0][2], r); }
  function fulfill(value) { resume("next", value); }
  function reject(value) { resume("throw", value); }
  function settle(f, v) { if (f(v), q.shift(), q.length) resume(q[0][0], q[0][1]); }
}

function __asyncDelegator(o) {
  var i, p;
  return i = {}, verb("next"), verb("throw", function (e) { throw e; }), verb("return"), i[Symbol.iterator] = function () { return this; }, i;
  function verb(n, f) { i[n] = o[n] ? function (v) { return (p = !p) ? { value: __await(o[n](v)), done: false } : f ? f(v) : v; } : f; }
}

function __asyncValues(o) {
  if (!Symbol.asyncIterator) throw new TypeError("Symbol.asyncIterator is not defined.");
  var m = o[Symbol.asyncIterator], i;
  return m ? m.call(o) : (o = typeof __values === "function" ? __values(o) : o[Symbol.iterator](), i = {}, verb("next"), verb("throw"), verb("return"), i[Symbol.asyncIterator] = function () { return this; }, i);
  function verb(n) { i[n] = o[n] && function (v) { return new Promise(function (resolve, reject) { v = o[n](v), settle(resolve, reject, v.done, v.value); }); }; }
  function settle(resolve, reject, d, v) { Promise.resolve(v).then(function(v) { resolve({ value: v, done: d }); }, reject); }
}

function __makeTemplateObject(cooked, raw) {
  if (Object.defineProperty) { Object.defineProperty(cooked, "raw", { value: raw }); } else { cooked.raw = raw; }
  return cooked;
};

var __setModuleDefault = Object.create ? (function(o, v) {
  Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
  o["default"] = v;
};

function __importStar(mod) {
  if (mod && mod.__esModule) return mod;
  var result = {};
  if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
  __setModuleDefault(result, mod);
  return result;
}

function __importDefault(mod) {
  return (mod && mod.__esModule) ? mod : { default: mod };
}

function __classPrivateFieldGet(receiver, state, kind, f) {
  if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
  if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
  return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
}

function __classPrivateFieldSet(receiver, state, value, kind, f) {
  if (kind === "m") throw new TypeError("Private method is not writable");
  if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a setter");
  if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot write private member to an object whose class did not declare it");
  return (kind === "a" ? f.call(receiver, value) : f ? f.value = value : state.set(receiver, value)), value;
}

function __classPrivateFieldIn(state, receiver) {
  if (receiver === null || (typeof receiver !== "object" && typeof receiver !== "function")) throw new TypeError("Cannot use 'in' operator on non-object");
  return typeof state === "function" ? receiver === state : state.has(receiver);
}

function __addDisposableResource(env, value, async) {
  if (value !== null && value !== void 0) {
    if (typeof value !== "object" && typeof value !== "function") throw new TypeError("Object expected.");
    var dispose, inner;
    if (async) {
      if (!Symbol.asyncDispose) throw new TypeError("Symbol.asyncDispose is not defined.");
      dispose = value[Symbol.asyncDispose];
    }
    if (dispose === void 0) {
      if (!Symbol.dispose) throw new TypeError("Symbol.dispose is not defined.");
      dispose = value[Symbol.dispose];
      if (async) inner = dispose;
    }
    if (typeof dispose !== "function") throw new TypeError("Object not disposable.");
    if (inner) dispose = function() { try { inner.call(this); } catch (e) { return Promise.reject(e); } };
    env.stack.push({ value: value, dispose: dispose, async: async });
  }
  else if (async) {
    env.stack.push({ async: true });
  }
  return value;
}

var _SuppressedError = typeof SuppressedError === "function" ? SuppressedError : function (error, suppressed, message) {
  var e = new Error(message);
  return e.name = "SuppressedError", e.error = error, e.suppressed = suppressed, e;
};

function __disposeResources(env) {
  function fail(e) {
    env.error = env.hasError ? new _SuppressedError(e, env.error, "An error was suppressed during disposal.") : e;
    env.hasError = true;
  }
  var r, s = 0;
  function next() {
    while (r = env.stack.pop()) {
      try {
        if (!r.async && s === 1) return s = 0, env.stack.push(r), Promise.resolve().then(next);
        if (r.dispose) {
          var result = r.dispose.call(r.value);
          if (r.async) return s |= 2, Promise.resolve(result).then(next, function(e) { fail(e); return next(); });
        }
        else s |= 1;
      }
      catch (e) {
        fail(e);
      }
    }
    if (s === 1) return env.hasError ? Promise.reject(env.error) : Promise.resolve();
    if (env.hasError) throw env.error;
  }
  return next();
}

/* unused harmony default export */ var __WEBPACK_DEFAULT_EXPORT__ = ({
  __extends,
  __assign,
  __rest,
  __decorate,
  __param,
  __metadata,
  __awaiter,
  __generator,
  __createBinding,
  __exportStar,
  __values,
  __read,
  __spread,
  __spreadArrays,
  __spreadArray,
  __await,
  __asyncGenerator,
  __asyncDelegator,
  __asyncValues,
  __makeTemplateObject,
  __importStar,
  __importDefault,
  __classPrivateFieldGet,
  __classPrivateFieldSet,
  __classPrivateFieldIn,
  __addDisposableResource,
  __disposeResources,
});


/***/ }),

/***/ 4346:
/*!***************************************************************!*\
  !*** ./node_modules/@pnp/sp/node_modules/tslib/tslib.es6.mjs ***!
  \***************************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   __decorate: () => (/* binding */ __decorate)
/* harmony export */ });
/* unused harmony exports __extends, __assign, __rest, __param, __esDecorate, __runInitializers, __propKey, __setFunctionName, __metadata, __awaiter, __generator, __createBinding, __exportStar, __values, __read, __spread, __spreadArrays, __spreadArray, __await, __asyncGenerator, __asyncDelegator, __asyncValues, __makeTemplateObject, __importStar, __importDefault, __classPrivateFieldGet, __classPrivateFieldSet, __classPrivateFieldIn, __addDisposableResource, __disposeResources */
/******************************************************************************
Copyright (c) Microsoft Corporation.

Permission to use, copy, modify, and/or distribute this software for any
purpose with or without fee is hereby granted.

THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES WITH
REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF MERCHANTABILITY
AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY SPECIAL, DIRECT,
INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES WHATSOEVER RESULTING FROM
LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR
OTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR
PERFORMANCE OF THIS SOFTWARE.
***************************************************************************** */
/* global Reflect, Promise, SuppressedError, Symbol, Iterator */

var extendStatics = function(d, b) {
  extendStatics = Object.setPrototypeOf ||
      ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
      function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
  return extendStatics(d, b);
};

function __extends(d, b) {
  if (typeof b !== "function" && b !== null)
      throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
  extendStatics(d, b);
  function __() { this.constructor = d; }
  d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
}

var __assign = function() {
  __assign = Object.assign || function __assign(t) {
      for (var s, i = 1, n = arguments.length; i < n; i++) {
          s = arguments[i];
          for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p)) t[p] = s[p];
      }
      return t;
  }
  return __assign.apply(this, arguments);
}

function __rest(s, e) {
  var t = {};
  for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p) && e.indexOf(p) < 0)
      t[p] = s[p];
  if (s != null && typeof Object.getOwnPropertySymbols === "function")
      for (var i = 0, p = Object.getOwnPropertySymbols(s); i < p.length; i++) {
          if (e.indexOf(p[i]) < 0 && Object.prototype.propertyIsEnumerable.call(s, p[i]))
              t[p[i]] = s[p[i]];
      }
  return t;
}

function __decorate(decorators, target, key, desc) {
  var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
  if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
  else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
  return c > 3 && r && Object.defineProperty(target, key, r), r;
}

function __param(paramIndex, decorator) {
  return function (target, key) { decorator(target, key, paramIndex); }
}

function __esDecorate(ctor, descriptorIn, decorators, contextIn, initializers, extraInitializers) {
  function accept(f) { if (f !== void 0 && typeof f !== "function") throw new TypeError("Function expected"); return f; }
  var kind = contextIn.kind, key = kind === "getter" ? "get" : kind === "setter" ? "set" : "value";
  var target = !descriptorIn && ctor ? contextIn["static"] ? ctor : ctor.prototype : null;
  var descriptor = descriptorIn || (target ? Object.getOwnPropertyDescriptor(target, contextIn.name) : {});
  var _, done = false;
  for (var i = decorators.length - 1; i >= 0; i--) {
      var context = {};
      for (var p in contextIn) context[p] = p === "access" ? {} : contextIn[p];
      for (var p in contextIn.access) context.access[p] = contextIn.access[p];
      context.addInitializer = function (f) { if (done) throw new TypeError("Cannot add initializers after decoration has completed"); extraInitializers.push(accept(f || null)); };
      var result = (0, decorators[i])(kind === "accessor" ? { get: descriptor.get, set: descriptor.set } : descriptor[key], context);
      if (kind === "accessor") {
          if (result === void 0) continue;
          if (result === null || typeof result !== "object") throw new TypeError("Object expected");
          if (_ = accept(result.get)) descriptor.get = _;
          if (_ = accept(result.set)) descriptor.set = _;
          if (_ = accept(result.init)) initializers.unshift(_);
      }
      else if (_ = accept(result)) {
          if (kind === "field") initializers.unshift(_);
          else descriptor[key] = _;
      }
  }
  if (target) Object.defineProperty(target, contextIn.name, descriptor);
  done = true;
};

function __runInitializers(thisArg, initializers, value) {
  var useValue = arguments.length > 2;
  for (var i = 0; i < initializers.length; i++) {
      value = useValue ? initializers[i].call(thisArg, value) : initializers[i].call(thisArg);
  }
  return useValue ? value : void 0;
};

function __propKey(x) {
  return typeof x === "symbol" ? x : "".concat(x);
};

function __setFunctionName(f, name, prefix) {
  if (typeof name === "symbol") name = name.description ? "[".concat(name.description, "]") : "";
  return Object.defineProperty(f, "name", { configurable: true, value: prefix ? "".concat(prefix, " ", name) : name });
};

function __metadata(metadataKey, metadataValue) {
  if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(metadataKey, metadataValue);
}

function __awaiter(thisArg, _arguments, P, generator) {
  function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
  return new (P || (P = Promise))(function (resolve, reject) {
      function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
      function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
      function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
      step((generator = generator.apply(thisArg, _arguments || [])).next());
  });
}

function __generator(thisArg, body) {
  var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
  return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
  function verb(n) { return function (v) { return step([n, v]); }; }
  function step(op) {
      if (f) throw new TypeError("Generator is already executing.");
      while (g && (g = 0, op[0] && (_ = 0)), _) try {
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
}

var __createBinding = Object.create ? (function(o, m, k, k2) {
  if (k2 === undefined) k2 = k;
  var desc = Object.getOwnPropertyDescriptor(m, k);
  if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
  }
  Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
  if (k2 === undefined) k2 = k;
  o[k2] = m[k];
});

function __exportStar(m, o) {
  for (var p in m) if (p !== "default" && !Object.prototype.hasOwnProperty.call(o, p)) __createBinding(o, m, p);
}

function __values(o) {
  var s = typeof Symbol === "function" && Symbol.iterator, m = s && o[s], i = 0;
  if (m) return m.call(o);
  if (o && typeof o.length === "number") return {
      next: function () {
          if (o && i >= o.length) o = void 0;
          return { value: o && o[i++], done: !o };
      }
  };
  throw new TypeError(s ? "Object is not iterable." : "Symbol.iterator is not defined.");
}

function __read(o, n) {
  var m = typeof Symbol === "function" && o[Symbol.iterator];
  if (!m) return o;
  var i = m.call(o), r, ar = [], e;
  try {
      while ((n === void 0 || n-- > 0) && !(r = i.next()).done) ar.push(r.value);
  }
  catch (error) { e = { error: error }; }
  finally {
      try {
          if (r && !r.done && (m = i["return"])) m.call(i);
      }
      finally { if (e) throw e.error; }
  }
  return ar;
}

/** @deprecated */
function __spread() {
  for (var ar = [], i = 0; i < arguments.length; i++)
      ar = ar.concat(__read(arguments[i]));
  return ar;
}

/** @deprecated */
function __spreadArrays() {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) s += arguments[i].length;
  for (var r = Array(s), k = 0, i = 0; i < il; i++)
      for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++)
          r[k] = a[j];
  return r;
}

function __spreadArray(to, from, pack) {
  if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
      if (ar || !(i in from)) {
          if (!ar) ar = Array.prototype.slice.call(from, 0, i);
          ar[i] = from[i];
      }
  }
  return to.concat(ar || Array.prototype.slice.call(from));
}

function __await(v) {
  return this instanceof __await ? (this.v = v, this) : new __await(v);
}

function __asyncGenerator(thisArg, _arguments, generator) {
  if (!Symbol.asyncIterator) throw new TypeError("Symbol.asyncIterator is not defined.");
  var g = generator.apply(thisArg, _arguments || []), i, q = [];
  return i = Object.create((typeof AsyncIterator === "function" ? AsyncIterator : Object).prototype), verb("next"), verb("throw"), verb("return", awaitReturn), i[Symbol.asyncIterator] = function () { return this; }, i;
  function awaitReturn(f) { return function (v) { return Promise.resolve(v).then(f, reject); }; }
  function verb(n, f) { if (g[n]) { i[n] = function (v) { return new Promise(function (a, b) { q.push([n, v, a, b]) > 1 || resume(n, v); }); }; if (f) i[n] = f(i[n]); } }
  function resume(n, v) { try { step(g[n](v)); } catch (e) { settle(q[0][3], e); } }
  function step(r) { r.value instanceof __await ? Promise.resolve(r.value.v).then(fulfill, reject) : settle(q[0][2], r); }
  function fulfill(value) { resume("next", value); }
  function reject(value) { resume("throw", value); }
  function settle(f, v) { if (f(v), q.shift(), q.length) resume(q[0][0], q[0][1]); }
}

function __asyncDelegator(o) {
  var i, p;
  return i = {}, verb("next"), verb("throw", function (e) { throw e; }), verb("return"), i[Symbol.iterator] = function () { return this; }, i;
  function verb(n, f) { i[n] = o[n] ? function (v) { return (p = !p) ? { value: __await(o[n](v)), done: false } : f ? f(v) : v; } : f; }
}

function __asyncValues(o) {
  if (!Symbol.asyncIterator) throw new TypeError("Symbol.asyncIterator is not defined.");
  var m = o[Symbol.asyncIterator], i;
  return m ? m.call(o) : (o = typeof __values === "function" ? __values(o) : o[Symbol.iterator](), i = {}, verb("next"), verb("throw"), verb("return"), i[Symbol.asyncIterator] = function () { return this; }, i);
  function verb(n) { i[n] = o[n] && function (v) { return new Promise(function (resolve, reject) { v = o[n](v), settle(resolve, reject, v.done, v.value); }); }; }
  function settle(resolve, reject, d, v) { Promise.resolve(v).then(function(v) { resolve({ value: v, done: d }); }, reject); }
}

function __makeTemplateObject(cooked, raw) {
  if (Object.defineProperty) { Object.defineProperty(cooked, "raw", { value: raw }); } else { cooked.raw = raw; }
  return cooked;
};

var __setModuleDefault = Object.create ? (function(o, v) {
  Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
  o["default"] = v;
};

function __importStar(mod) {
  if (mod && mod.__esModule) return mod;
  var result = {};
  if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
  __setModuleDefault(result, mod);
  return result;
}

function __importDefault(mod) {
  return (mod && mod.__esModule) ? mod : { default: mod };
}

function __classPrivateFieldGet(receiver, state, kind, f) {
  if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
  if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
  return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
}

function __classPrivateFieldSet(receiver, state, value, kind, f) {
  if (kind === "m") throw new TypeError("Private method is not writable");
  if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a setter");
  if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot write private member to an object whose class did not declare it");
  return (kind === "a" ? f.call(receiver, value) : f ? f.value = value : state.set(receiver, value)), value;
}

function __classPrivateFieldIn(state, receiver) {
  if (receiver === null || (typeof receiver !== "object" && typeof receiver !== "function")) throw new TypeError("Cannot use 'in' operator on non-object");
  return typeof state === "function" ? receiver === state : state.has(receiver);
}

function __addDisposableResource(env, value, async) {
  if (value !== null && value !== void 0) {
    if (typeof value !== "object" && typeof value !== "function") throw new TypeError("Object expected.");
    var dispose, inner;
    if (async) {
      if (!Symbol.asyncDispose) throw new TypeError("Symbol.asyncDispose is not defined.");
      dispose = value[Symbol.asyncDispose];
    }
    if (dispose === void 0) {
      if (!Symbol.dispose) throw new TypeError("Symbol.dispose is not defined.");
      dispose = value[Symbol.dispose];
      if (async) inner = dispose;
    }
    if (typeof dispose !== "function") throw new TypeError("Object not disposable.");
    if (inner) dispose = function() { try { inner.call(this); } catch (e) { return Promise.reject(e); } };
    env.stack.push({ value: value, dispose: dispose, async: async });
  }
  else if (async) {
    env.stack.push({ async: true });
  }
  return value;
}

var _SuppressedError = typeof SuppressedError === "function" ? SuppressedError : function (error, suppressed, message) {
  var e = new Error(message);
  return e.name = "SuppressedError", e.error = error, e.suppressed = suppressed, e;
};

function __disposeResources(env) {
  function fail(e) {
    env.error = env.hasError ? new _SuppressedError(e, env.error, "An error was suppressed during disposal.") : e;
    env.hasError = true;
  }
  var r, s = 0;
  function next() {
    while (r = env.stack.pop()) {
      try {
        if (!r.async && s === 1) return s = 0, env.stack.push(r), Promise.resolve().then(next);
        if (r.dispose) {
          var result = r.dispose.call(r.value);
          if (r.async) return s |= 2, Promise.resolve(result).then(next, function(e) { fail(e); return next(); });
        }
        else s |= 1;
      }
      catch (e) {
        fail(e);
      }
    }
    if (s === 1) return env.hasError ? Promise.reject(env.error) : Promise.resolve();
    if (env.hasError) throw env.error;
  }
  return next();
}

/* unused harmony default export */ var __WEBPACK_DEFAULT_EXPORT__ = ({
  __extends,
  __assign,
  __rest,
  __decorate,
  __param,
  __metadata,
  __awaiter,
  __generator,
  __createBinding,
  __exportStar,
  __values,
  __read,
  __spread,
  __spreadArrays,
  __spreadArray,
  __await,
  __asyncGenerator,
  __asyncDelegator,
  __asyncValues,
  __makeTemplateObject,
  __importStar,
  __importDefault,
  __classPrivateFieldGet,
  __classPrivateFieldSet,
  __classPrivateFieldIn,
  __addDisposableResource,
  __disposeResources,
});


/***/ }),

/***/ 7945:
/*!*********************************************************!*\
  !*** ./node_modules/@pnp/core/behaviors/assign-from.js ***!
  \*********************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   AssignFrom: () => (/* binding */ AssignFrom)
/* harmony export */ });
/**
 * Behavior that will assign a ref to the source's observers and reset the instance's inheriting flag
 *
 * @param source The source instance from which we will assign the observers
 */
function AssignFrom(source) {
    return (instance) => {
        instance.observers = source.observers;
        instance._inheritingObservers = true;
        return instance;
    };
}


/***/ }),

/***/ 4337:
/*!*******************************************************!*\
  !*** ./node_modules/@pnp/core/behaviors/copy-from.js ***!
  \*******************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   CopyFrom: () => (/* binding */ CopyFrom)
/* harmony export */ });
/* harmony import */ var _util_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../util.js */ 2165);
/* harmony import */ var _timeline_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../timeline.js */ 447);


/**
 * Behavior that will copy all the observers in the source timeline and apply it to the incoming instance
 *
 * @param source The source instance from which we will copy the observers
 * @param behavior replace = observers are cleared before adding, append preserves any observers already present
 * @param filter If provided filters the moments from which the observers are copied. It should return true for each moment to include.
 * @returns The mutated this
 */
function CopyFrom(source, behavior = "append", filter) {
    return (instance) => {
        return Reflect.apply(copyObservers, instance, [source, behavior, filter]);
    };
}
/**
 * Function with implied this allows us to access protected members
 *
 * @param this The timeline whose observers we will copy
 * @param source The source instance from which we will copy the observers
 * @param behavior replace = observers are cleared before adding, append preserves any observers already present
 * @returns The mutated this
 */
function copyObservers(source, behavior, filter) {
    if (!(0,_util_js__WEBPACK_IMPORTED_MODULE_1__.objectDefinedNotNull)(source) || !(0,_util_js__WEBPACK_IMPORTED_MODULE_1__.objectDefinedNotNull)(source.observers)) {
        return this;
    }
    if (!(0,_util_js__WEBPACK_IMPORTED_MODULE_1__.isFunc)(filter)) {
        filter = () => true;
    }
    const clonedSource = (0,_timeline_js__WEBPACK_IMPORTED_MODULE_0__.cloneObserverCollection)(source.observers);
    const keys = Object.keys(clonedSource).filter(filter);
    for (let i = 0; i < keys.length; i++) {
        const key = keys[i];
        const on = this.on[key];
        if (behavior === "replace") {
            on.clear();
        }
        const momentObservers = clonedSource[key];
        momentObservers.forEach(v => on(v));
    }
    return this;
}


/***/ }),

/***/ 1971:
/*!*****************************************!*\
  !*** ./node_modules/@pnp/core/index.js ***!
  \*****************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   AssignFrom: () => (/* reexport safe */ _behaviors_assign_from_js__WEBPACK_IMPORTED_MODULE_4__.AssignFrom),
/* harmony export */   CopyFrom: () => (/* reexport safe */ _behaviors_copy_from_js__WEBPACK_IMPORTED_MODULE_5__.CopyFrom),
/* harmony export */   PnPClientStorage: () => (/* reexport safe */ _storage_js__WEBPACK_IMPORTED_MODULE_0__.PnPClientStorage),
/* harmony export */   Timeline: () => (/* reexport safe */ _timeline_js__WEBPACK_IMPORTED_MODULE_3__.Timeline),
/* harmony export */   asyncBroadcast: () => (/* reexport safe */ _moments_js__WEBPACK_IMPORTED_MODULE_2__.asyncBroadcast),
/* harmony export */   asyncReduce: () => (/* reexport safe */ _moments_js__WEBPACK_IMPORTED_MODULE_2__.asyncReduce),
/* harmony export */   broadcast: () => (/* reexport safe */ _moments_js__WEBPACK_IMPORTED_MODULE_2__.broadcast),
/* harmony export */   combine: () => (/* reexport safe */ _util_js__WEBPACK_IMPORTED_MODULE_1__.combine),
/* harmony export */   dateAdd: () => (/* reexport safe */ _util_js__WEBPACK_IMPORTED_MODULE_1__.dateAdd),
/* harmony export */   delay: () => (/* reexport safe */ _util_js__WEBPACK_IMPORTED_MODULE_1__.delay),
/* harmony export */   getGUID: () => (/* reexport safe */ _util_js__WEBPACK_IMPORTED_MODULE_1__.getGUID),
/* harmony export */   getHashCode: () => (/* reexport safe */ _util_js__WEBPACK_IMPORTED_MODULE_1__.getHashCode),
/* harmony export */   hOP: () => (/* reexport safe */ _util_js__WEBPACK_IMPORTED_MODULE_1__.hOP),
/* harmony export */   isArray: () => (/* reexport safe */ _util_js__WEBPACK_IMPORTED_MODULE_1__.isArray),
/* harmony export */   isFunc: () => (/* reexport safe */ _util_js__WEBPACK_IMPORTED_MODULE_1__.isFunc),
/* harmony export */   isUrlAbsolute: () => (/* reexport safe */ _util_js__WEBPACK_IMPORTED_MODULE_1__.isUrlAbsolute),
/* harmony export */   jsS: () => (/* reexport safe */ _util_js__WEBPACK_IMPORTED_MODULE_1__.jsS),
/* harmony export */   lifecycle: () => (/* reexport safe */ _moments_js__WEBPACK_IMPORTED_MODULE_2__.lifecycle),
/* harmony export */   noInherit: () => (/* reexport safe */ _timeline_js__WEBPACK_IMPORTED_MODULE_3__.noInherit),
/* harmony export */   objectDefinedNotNull: () => (/* reexport safe */ _util_js__WEBPACK_IMPORTED_MODULE_1__.objectDefinedNotNull),
/* harmony export */   reduce: () => (/* reexport safe */ _moments_js__WEBPACK_IMPORTED_MODULE_2__.reduce),
/* harmony export */   request: () => (/* reexport safe */ _moments_js__WEBPACK_IMPORTED_MODULE_2__.request),
/* harmony export */   stringIsNullOrEmpty: () => (/* reexport safe */ _util_js__WEBPACK_IMPORTED_MODULE_1__.stringIsNullOrEmpty)
/* harmony export */ });
/* harmony import */ var _storage_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./storage.js */ 1922);
/* harmony import */ var _util_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./util.js */ 2165);
/* harmony import */ var _moments_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./moments.js */ 7955);
/* harmony import */ var _timeline_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./timeline.js */ 447);
/* harmony import */ var _behaviors_assign_from_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./behaviors/assign-from.js */ 7945);
/* harmony import */ var _behaviors_copy_from_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./behaviors/copy-from.js */ 4337);




/**
 * Behavior exports
 */




/***/ }),

/***/ 7955:
/*!*******************************************!*\
  !*** ./node_modules/@pnp/core/moments.js ***!
  \*******************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   asyncBroadcast: () => (/* binding */ asyncBroadcast),
/* harmony export */   asyncReduce: () => (/* binding */ asyncReduce),
/* harmony export */   broadcast: () => (/* binding */ broadcast),
/* harmony export */   lifecycle: () => (/* binding */ lifecycle),
/* harmony export */   reduce: () => (/* binding */ reduce),
/* harmony export */   request: () => (/* binding */ request)
/* harmony export */ });
/* harmony import */ var _util_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./util.js */ 2165);

/**
 * Emits to all registered observers the supplied arguments. Any values returned by the observers are ignored
 *
 * @returns void
 */
function broadcast() {
    return function (observers, ...args) {
        const obs = [...observers];
        for (let i = 0; i < obs.length; i++) {
            Reflect.apply(obs[i], this, args);
        }
    };
}
/**
 * Defines a moment that executes each observer asynchronously in parallel awaiting all promises to resolve or reject before continuing
 *
 * @returns The final set of arguments
 */
function asyncBroadcast() {
    return async function (observers, ...args) {
        // get our initial values
        const r = args;
        const obs = [...observers];
        const promises = [];
        for (let i = 0; i < obs.length; i++) {
            promises.push(Reflect.apply(obs[i], this, r));
        }
        return Promise.all(promises);
    };
}
/**
 * Defines a moment that executes each observer synchronously, passing the returned arguments as the arguments to the next observer.
 * This is very much like the redux pattern taking the arguments as the state which each observer may modify then returning a new state
 *
 * @returns The final set of arguments
 */
function reduce() {
    return function (observers, ...args) {
        const obs = [...observers];
        return obs.reduce((params, func) => Reflect.apply(func, this, params), args);
    };
}
/**
 * Defines a moment that executes each observer asynchronously, awaiting the result and passes the returned arguments as the arguments to the next observer.
 * This is very much like the redux pattern taking the arguments as the state which each observer may modify then returning a new state
 *
 * @returns The final set of arguments
 */
function asyncReduce() {
    return async function (observers, ...args) {
        const obs = [...observers];
        return obs.reduce((prom, func) => prom.then((params) => Reflect.apply(func, this, params)), Promise.resolve(args));
    };
}
/**
 * Defines a moment where the first registered observer is used to asynchronously execute a request, returning a single result
 * If no result is returned (undefined) no further action is taken and the result will be undefined (i.e. additional observers are not used)
 *
 * @returns The result returned by the first registered observer
 */
function request() {
    return async function (observers, ...args) {
        if (!(0,_util_js__WEBPACK_IMPORTED_MODULE_0__.isArray)(observers) || observers.length < 1) {
            return undefined;
        }
        const handler = observers[0];
        return Reflect.apply(handler, this, args);
    };
}
/**
 * Defines a special moment used to configure the timeline itself before starting. Each observer is executed in order,
 * possibly modifying the "this" instance, with the final product returned
 *
 */
function lifecycle() {
    return function (observers, ...args) {
        const obs = [...observers];
        // process each handler which updates our instance in order
        // very similar to asyncReduce but the state is the object itself
        for (let i = 0; i < obs.length; i++) {
            Reflect.apply(obs[i], this, args);
        }
        return this;
    };
}


/***/ }),

/***/ 1922:
/*!*******************************************!*\
  !*** ./node_modules/@pnp/core/storage.js ***!
  \*******************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   PnPClientStorage: () => (/* binding */ PnPClientStorage)
/* harmony export */ });
/* unused harmony export PnPClientStorageWrapper */
/* harmony import */ var _util_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./util.js */ 2165);

let storageShim;
function getStorageShim() {
    if (typeof storageShim === "undefined") {
        storageShim = new MemoryStorage();
    }
    return storageShim;
}
/**
 * A wrapper class to provide a consistent interface to browser based storage
 *
 */
class PnPClientStorageWrapper {
    /**
     * Creates a new instance of the PnPClientStorageWrapper class
     *
     * @constructor
     */
    constructor(store) {
        this.store = store;
        this.enabled = this.test();
    }
    /**
     * Get a value from storage, or null if that value does not exist
     *
     * @param key The key whose value we want to retrieve
     */
    get(key) {
        if (!this.enabled) {
            return null;
        }
        const o = this.store.getItem(key);
        if (!(0,_util_js__WEBPACK_IMPORTED_MODULE_0__.objectDefinedNotNull)(o)) {
            return null;
        }
        const persistable = JSON.parse(o);
        if (new Date(persistable.expiration) <= new Date()) {
            this.delete(key);
            return null;
        }
        else {
            return persistable.value;
        }
    }
    /**
     * Adds a value to the underlying storage
     *
     * @param key The key to use when storing the provided value
     * @param o The value to store
     * @param expire Optional, if provided the expiration of the item, otherwise the default is used
     */
    put(key, o, expire) {
        if (this.enabled) {
            this.store.setItem(key, this.createPersistable(o, expire));
        }
    }
    /**
     * Deletes a value from the underlying storage
     *
     * @param key The key of the pair we want to remove from storage
     */
    delete(key) {
        if (this.enabled) {
            this.store.removeItem(key);
        }
    }
    /**
     * Gets an item from the underlying storage, or adds it if it does not exist using the supplied getter function
     *
     * @param key The key to use when storing the provided value
     * @param getter A function which will upon execution provide the desired value
     * @param expire Optional, if provided the expiration of the item, otherwise the default is used
     */
    async getOrPut(key, getter, expire) {
        if (!this.enabled) {
            return getter();
        }
        let o = this.get(key);
        if (o === null) {
            o = await getter();
            this.put(key, o, expire);
        }
        return o;
    }
    /**
     * Deletes any expired items placed in the store by the pnp library, leaves other items untouched
     */
    async deleteExpired() {
        if (!this.enabled) {
            return;
        }
        for (let i = 0; i < this.store.length; i++) {
            const key = this.store.key(i);
            if (key !== null) {
                // test the stored item to see if we stored it
                if (/["|']?pnp["|']? ?: ?1/i.test(this.store.getItem(key))) {
                    // get those items as get will delete from cache if they are expired
                    await this.get(key);
                }
            }
        }
    }
    /**
     * Used to determine if the wrapped storage is available currently
     */
    test() {
        const str = "t";
        try {
            this.store.setItem(str, str);
            this.store.removeItem(str);
            return true;
        }
        catch (e) {
            return false;
        }
    }
    /**
     * Creates the persistable to store
     */
    createPersistable(o, expire) {
        if (expire === undefined) {
            expire = (0,_util_js__WEBPACK_IMPORTED_MODULE_0__.dateAdd)(new Date(), "minute", 5);
        }
        return (0,_util_js__WEBPACK_IMPORTED_MODULE_0__.jsS)({ pnp: 1, expiration: expire, value: o });
    }
}
/**
 * A thin implementation of in-memory storage for use in nodejs
 */
class MemoryStorage {
    constructor(_store = new Map()) {
        this._store = _store;
    }
    get length() {
        return this._store.size;
    }
    clear() {
        this._store.clear();
    }
    getItem(key) {
        return this._store.get(key);
    }
    key(index) {
        return Array.from(this._store)[index][0];
    }
    removeItem(key) {
        this._store.delete(key);
    }
    setItem(key, data) {
        this._store.set(key, data);
    }
}
/**
 * A class that will establish wrappers for both local and session storage, substituting basic memory storage for nodejs
 */
class PnPClientStorage {
    /**
     * Creates a new instance of the PnPClientStorage class
     *
     * @constructor
     */
    constructor(_local = null, _session = null) {
        this._local = _local;
        this._session = _session;
    }
    /**
     * Provides access to the local storage of the browser
     */
    get local() {
        if (this._local === null) {
            this._local = new PnPClientStorageWrapper(typeof localStorage === "undefined" ? getStorageShim() : localStorage);
        }
        return this._local;
    }
    /**
     * Provides access to the session storage of the browser
     */
    get session() {
        if (this._session === null) {
            this._session = new PnPClientStorageWrapper(typeof sessionStorage === "undefined" ? getStorageShim() : sessionStorage);
        }
        return this._session;
    }
}


/***/ }),

/***/ 447:
/*!********************************************!*\
  !*** ./node_modules/@pnp/core/timeline.js ***!
  \********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Timeline: () => (/* binding */ Timeline),
/* harmony export */   cloneObserverCollection: () => (/* binding */ cloneObserverCollection),
/* harmony export */   noInherit: () => (/* binding */ noInherit)
/* harmony export */ });
/* unused harmony export once */
/* harmony import */ var _moments_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./moments.js */ 7955);
/* harmony import */ var _util_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./util.js */ 2165);


/**
 * Field name to hold any flags on observer functions used to modify their behavior
 */
const flags = Symbol.for("ObserverLifecycleFlags");
/**
 * Creates a filter function for use in Array.filter that will filter OUT any observers with the specified [flag]
 *
 * @param flag The flag used to exclude observers
 * @returns An Array.filter function
 */
// eslint-disable-next-line no-bitwise
const byFlag = (flag) => ((observer) => !((observer[flags] || 0) & flag));
/**
 * Creates an observer lifecycle modification flag application function
 * @param flag The flag to the bound function should add
 * @returns A function that can be used to apply [flag] to any valid observer
 */
const addFlag = (flag) => ((observer) => {
    // eslint-disable-next-line no-bitwise
    observer[flags] = (observer[flags] || 0) | flag;
    return observer;
});
/**
 * Observer lifecycle modifier that indicates this observer should NOT be inherited by any child
 * timelines.
 */
const noInherit = addFlag(1 /* ObserverLifecycleFlags.noInherit */);
/**
 * Observer lifecycle modifier that indicates this observer should only fire once per instance, it is then removed.
 *
 * Note: If you have a parent and child timeline "once" will affect both and the observer will fire once for a parent lifecycle
 * and once for a child lifecycle
 */
const once = addFlag(2 /* ObserverLifecycleFlags.once */);
/**
 * Timeline represents a set of operations executed in order of definition,
 * with each moment's behavior controlled by the implementing function
 */
class Timeline {
    /**
     * Creates a new instance of Timeline with the supplied moments and optionally any observers to include
     *
     * @param moments The moment object defining this timeline
     * @param observers Any observers to include (optional)
     */
    constructor(moments, observers = {}) {
        this.moments = moments;
        this.observers = observers;
        this._onProxy = null;
        this._emitProxy = null;
        this._inheritingObservers = true;
    }
    /**
     * Apply the supplied behavior(s) to this timeline
     *
     * @param behaviors One or more behaviors
     * @returns `this` Timeline
     */
    using(...behaviors) {
        for (let i = 0; i < behaviors.length; i++) {
            behaviors[i](this);
        }
        return this;
    }
    /**
     * Property allowing access to manage observers on moments within this timeline
     */
    get on() {
        if (this._onProxy === null) {
            this._onProxy = new Proxy(this, {
                get: (target, p) => Object.assign((handler) => {
                    target.cloneObserversOnChange();
                    addObserver(target.observers, p, handler, 1 /* ObserverAddBehavior.Add */);
                    return target;
                }, {
                    toArray: () => {
                        return Reflect.has(target.observers, p) ? [...Reflect.get(target.observers, p)] : [];
                    },
                    replace: (handler) => {
                        target.cloneObserversOnChange();
                        addObserver(target.observers, p, handler, 3 /* ObserverAddBehavior.Replace */);
                        return target;
                    },
                    prepend: (handler) => {
                        target.cloneObserversOnChange();
                        addObserver(target.observers, p, handler, 2 /* ObserverAddBehavior.Prepend */);
                        return target;
                    },
                    clear: () => {
                        if (Reflect.has(target.observers, p)) {
                            target.cloneObserversOnChange();
                            // we trust ourselves that this will be an array
                            target.observers[p].length = 0;
                            return true;
                        }
                        return false;
                    },
                }),
            });
        }
        return this._onProxy;
    }
    /**
     * Shorthand method to emit a logging event tied to this timeline
     *
     * @param message The message to log
     * @param level The level at which the message applies
     */
    log(message, level = 0) {
        this.emit.log(message, level);
    }
    /**
     * Shorthand method to emit an error event tied to this timeline
     *
     * @param e Optional. Any error object to emit. If none is provided no emit occurs
     */
    error(e) {
        if ((0,_util_js__WEBPACK_IMPORTED_MODULE_0__.objectDefinedNotNull)(e)) {
            this.emit.error(e);
        }
    }
    /**
     * Property allowing access to invoke a moment from within this timeline
     */
    get emit() {
        if (this._emitProxy === null) {
            this._emitProxy = new Proxy(this, {
                get: (target, p) => (...args) => {
                    // handle the case where no observers registered for the target moment
                    const observers = Reflect.has(target.observers, p) ? Reflect.get(target.observers, p) : [];
                    if ((!(0,_util_js__WEBPACK_IMPORTED_MODULE_0__.isArray)(observers) || observers.length < 1) && p === "error") {
                        // if we are emitting an error, and no error observers are defined, we throw
                        throw Error(`Unhandled Exception: ${args[0]}`);
                    }
                    try {
                        // default to broadcasting any events without specific impl (will apply to log and error)
                        const moment = Reflect.has(target.moments, p) ? Reflect.get(target.moments, p) : p === "init" || p === "dispose" ? (0,_moments_js__WEBPACK_IMPORTED_MODULE_1__.lifecycle)() : (0,_moments_js__WEBPACK_IMPORTED_MODULE_1__.broadcast)();
                        // pass control to the individual moment's implementation
                        return Reflect.apply(moment, target, [observers, ...args]);
                    }
                    catch (e) {
                        if (p !== "error") {
                            this.error(e);
                        }
                        else {
                            // if all else fails, re-throw as we are getting errors from error observers meaning something is sideways
                            throw e;
                        }
                    }
                    finally {
                        // here we need to remove any "once" observers
                        if (observers && observers.length > 0) {
                            Reflect.set(target.observers, p, observers.filter(byFlag(2 /* ObserverLifecycleFlags.once */)));
                        }
                    }
                },
            });
        }
        return this._emitProxy;
    }
    /**
     * Starts a timeline
     *
     * @description This method first emits "init" to allow for any needed initial conditions then calls execute with any supplied init
     *
     * @param init A value passed into the execute logic from the initiator of the timeline
     * @returns The result of this.execute
     */
    start(init) {
        // initialize our timeline
        this.emit.init();
        // get a ref to the promise returned by execute
        const p = this.execute(init);
        // attach our dispose logic
        p.finally(() => {
            try {
                // provide an opportunity for cleanup of the timeline
                this.emit.dispose();
            }
            catch (e) {
                // shouldn't happen, but possible dispose throws - which may be missed as the usercode await will have resolved.
                const e2 = Object.assign(Error("Error in dispose."), { innerException: e });
                this.error(e2);
            }
        }).catch(() => void (0));
        // give the promise back to the caller
        return p;
    }
    /**
     * By default a timeline references the same observer collection as a parent timeline,
     * if any changes are made to the observers this method first clones them ensuring we
     * maintain a local copy and de-ref the parent
     */
    cloneObserversOnChange() {
        if (this._inheritingObservers) {
            this._inheritingObservers = false;
            this.observers = cloneObserverCollection(this.observers);
        }
    }
}
/**
 * Adds an observer to a given target
 *
 * @param target The object to which events are registered
 * @param moment The name of the moment to which the observer is registered
 * @param addBehavior Determines how the observer is added to the collection
 *
 */
function addObserver(target, moment, observer, addBehavior) {
    if (!(0,_util_js__WEBPACK_IMPORTED_MODULE_0__.isFunc)(observer)) {
        throw Error("Observers must be functions.");
    }
    if (!Reflect.has(target, moment)) {
        // if we don't have a registration for this moment, then we just add a new prop
        target[moment] = [observer];
    }
    else {
        // if we have an existing property then we follow the specified behavior
        switch (addBehavior) {
            case 1 /* ObserverAddBehavior.Add */:
                target[moment].push(observer);
                break;
            case 2 /* ObserverAddBehavior.Prepend */:
                target[moment].unshift(observer);
                break;
            case 3 /* ObserverAddBehavior.Replace */:
                target[moment].length = 0;
                target[moment].push(observer);
                break;
        }
    }
    return target[moment];
}
function cloneObserverCollection(source) {
    return Reflect.ownKeys(source).reduce((clone, key) => {
        clone[key] = [...source[key].filter(byFlag(1 /* ObserverLifecycleFlags.noInherit */))];
        return clone;
    }, {});
}


/***/ }),

/***/ 2165:
/*!****************************************!*\
  !*** ./node_modules/@pnp/core/util.js ***!
  \****************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   combine: () => (/* binding */ combine),
/* harmony export */   dateAdd: () => (/* binding */ dateAdd),
/* harmony export */   delay: () => (/* binding */ delay),
/* harmony export */   getGUID: () => (/* binding */ getGUID),
/* harmony export */   getHashCode: () => (/* binding */ getHashCode),
/* harmony export */   hOP: () => (/* binding */ hOP),
/* harmony export */   isArray: () => (/* binding */ isArray),
/* harmony export */   isFunc: () => (/* binding */ isFunc),
/* harmony export */   isUrlAbsolute: () => (/* binding */ isUrlAbsolute),
/* harmony export */   jsS: () => (/* binding */ jsS),
/* harmony export */   objectDefinedNotNull: () => (/* binding */ objectDefinedNotNull),
/* harmony export */   stringIsNullOrEmpty: () => (/* binding */ stringIsNullOrEmpty)
/* harmony export */ });
/* unused harmony exports getRandomString, parseToAtob */
/**
 * Adds a value to a date
 *
 * @param date The date to which we will add units, done in local time
 * @param interval The name of the interval to add, one of: ['year', 'quarter', 'month', 'week', 'day', 'hour', 'minute', 'second']
 * @param units The amount to add to date of the given interval
 *
 * http://stackoverflow.com/questions/1197928/how-to-add-30-minutes-to-a-javascript-date-object
 */
function dateAdd(date, interval, units) {
    let ret = new Date(date.toString()); // don't change original date
    switch (interval.toLowerCase()) {
        case "year":
            ret.setFullYear(ret.getFullYear() + units);
            break;
        case "quarter":
            ret.setMonth(ret.getMonth() + 3 * units);
            break;
        case "month":
            ret.setMonth(ret.getMonth() + units);
            break;
        case "week":
            ret.setDate(ret.getDate() + 7 * units);
            break;
        case "day":
            ret.setDate(ret.getDate() + units);
            break;
        case "hour":
            ret.setTime(ret.getTime() + units * 3600000);
            break;
        case "minute":
            ret.setTime(ret.getTime() + units * 60000);
            break;
        case "second":
            ret.setTime(ret.getTime() + units * 1000);
            break;
        default:
            ret = undefined;
            break;
    }
    return ret;
}
/**
 * Combines an arbitrary set of paths ensuring and normalizes the slashes
 *
 * @param paths 0 to n path parts to combine
 */
function combine(...paths) {
    return paths
        .filter(path => !stringIsNullOrEmpty(path))
        .map(path => path.replace(/^[\\|/]/, "").replace(/[\\|/]$/, ""))
        .join("/")
        .replace(/\\/g, "/");
}
/**
 * Gets a random string of chars length
 *
 * https://stackoverflow.com/questions/1349404/generate-random-string-characters-in-javascript
 *
 * @param chars The length of the random string to generate
 */
function getRandomString(chars) {
    const text = new Array(chars);
    for (let i = 0; i < chars; i++) {
        text[i] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789".charAt(Math.floor(Math.random() * 62));
    }
    return text.join("");
}
/**
 * Gets a random GUID value
 *
 * http://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript
 */
/* eslint-disable no-bitwise */
function getGUID() {
    let d = Date.now();
    return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, function (c) {
        const r = (d + Math.random() * 16) % 16 | 0;
        d = Math.floor(d / 16);
        return (c === "x" ? r : (r & 0x3 | 0x8)).toString(16);
    });
}
/* eslint-enable no-bitwise */
/**
 * Determines if a given value is a function
 *
 * @param f The thing to test for functionness
 */
// eslint-disable-next-line @typescript-eslint/ban-types
function isFunc(f) {
    return typeof f === "function";
}
/**
 * @returns whether the provided parameter is a JavaScript Array or not.
*/
function isArray(array) {
    return Array.isArray(array);
}
/**
 * Determines if a given url is absolute
 *
 * @param url The url to check to see if it is absolute
 */
function isUrlAbsolute(url) {
    return /^https?:\/\/|^\/\//i.test(url);
}
/**
 * Determines if a string is null or empty or undefined
 *
 * @param s The string to test
 */
function stringIsNullOrEmpty(s) {
    return typeof s === "undefined" || s === null || s.length < 1;
}
/**
 * Determines if an object is both defined and not null
 * @param obj Object to test
 */
function objectDefinedNotNull(obj) {
    return typeof obj !== "undefined" && obj !== null;
}
/**
 * Shorthand for JSON.stringify
 *
 * @param o Any type of object
 */
function jsS(o) {
    return JSON.stringify(o);
}
/**
 * Shorthand for Object.hasOwnProperty
 *
 * @param o Object to check for
 * @param p Name of the property
 */
function hOP(o, p) {
    return Object.hasOwnProperty.call(o, p);
}
/**
 * @returns validates and returns a valid atob conversion
*/
function parseToAtob(str) {
    const base64Regex = /^[A-Za-z0-9+/]+={0,2}$/;
    try {
        // test if str has been JSON.stringified
        const parsed = JSON.parse(str);
        if (base64Regex.test(parsed)) {
            return atob(parsed);
        }
        return null;
    }
    catch (err) {
        // Not a valid JSON string, check if it's a standalone Base64 string
        return base64Regex.test(str) ? atob(str) : null;
    }
}
/**
 * Generates a ~unique hash code
 *
 * From: https://stackoverflow.com/questions/6122571/simple-non-secure-hash-function-for-javascript
 */
/* eslint-disable no-bitwise */
function getHashCode(s) {
    let hash = 0;
    if (s.length === 0) {
        return hash;
    }
    for (let i = 0; i < s.length; i++) {
        const chr = s.charCodeAt(i);
        hash = ((hash << 5) - hash) + chr;
        hash |= 0; // Convert to 32bit integer
    }
    return hash;
}
/* eslint-enable no-bitwise */
/**
 * Waits a specified number of milliseconds before resolving
 *
 * @param ms Number of ms to wait
 */
function delay(ms) {
    return new Promise((resolve) => {
        setTimeout(resolve, ms);
    });
}


/***/ }),

/***/ 8786:
/*!****************************************************************!*\
  !*** ./node_modules/@pnp/queryable/behaviors/browser-fetch.js ***!
  \****************************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   BrowserFetchWithRetry: () => (/* binding */ BrowserFetchWithRetry)
/* harmony export */ });
/* unused harmony export BrowserFetch */
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _parsers_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./parsers.js */ 5102);


function BrowserFetch(props) {
    const { replace } = {
        replace: true,
        ...props,
    };
    return (instance) => {
        if (replace) {
            instance.on.send.clear();
        }
        instance.on.send(function (url, init) {
            this.log(`Fetch: ${init.method} ${url.toString()}`, 0);
            return fetch(url.toString(), init);
        });
        return instance;
    };
}
function BrowserFetchWithRetry(props) {
    const { interval, replace, retries } = {
        replace: true,
        interval: 200,
        retries: 3,
        ...props,
    };
    return (instance) => {
        if (replace) {
            instance.on.send.clear();
        }
        instance.on.send(function (url, init) {
            let response;
            let wait = interval;
            let count = 0;
            let lastErr;
            const retry = async () => {
                // if we've tried too many times, throw
                if (count >= retries) {
                    throw lastErr || new _parsers_js__WEBPACK_IMPORTED_MODULE_1__.HttpRequestError(`Retry count exceeded (${retries}) for this request. ${response.status}: ${response.statusText};`, response);
                }
                count++;
                if (typeof response === "undefined" || (response === null || response === void 0 ? void 0 : response.status) === 429 || (response === null || response === void 0 ? void 0 : response.status) === 503 || (response === null || response === void 0 ? void 0 : response.status) === 504) {
                    // this is our first try and response isn't defined yet
                    // we have been throttled OR http status code 503 or 504, we can retry this
                    if (typeof response !== "undefined") {
                        // this isn't our first try so we need to calculate delay
                        if (response.headers.has("Retry-After")) {
                            // if we have gotten a header, use that value as the delay value in seconds
                            // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
                            wait = parseInt(response.headers.get("Retry-After"), 10) * 1000;
                        }
                        else {
                            // Increment our counters.
                            wait *= 2;
                        }
                        this.log(`Attempt #${count} to retry request which failed with ${response.status}: ${response.statusText}`, 0);
                        await (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.delay)(wait);
                    }
                    try {
                        const u = url.toString();
                        this.log(`Fetch: ${init.method} ${u}`, 0);
                        response = await fetch(u, init);
                        // if we got a good response, return it, otherwise see if we can retry
                        return response.ok ? response : retry();
                    }
                    catch (err) {
                        if (/AbortError/.test(err.name)) {
                            // don't retry aborted requests
                            throw err;
                        }
                        // if there is no network the response is undefined and err is all we have
                        // so we grab the err and save it to throw if we exceed the number of retries
                        // #2226 first reported this
                        lastErr = err;
                        return retry();
                    }
                }
                else {
                    return response;
                }
            };
            // this the the first call to retry that starts the cycle
            // response is undefined and the other values have their defaults
            return retry();
        });
        return instance;
    };
}


/***/ }),

/***/ 7395:
/*!**********************************************************************!*\
  !*** ./node_modules/@pnp/queryable/behaviors/caching-pessimistic.js ***!
  \**********************************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* unused harmony export CachingPessimisticRefresh */
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _queryable_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../queryable.js */ 7111);
/* harmony import */ var _caching_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./caching.js */ 6936);



/**
 * Pessimistic Caching Behavior
 * Always returns the cached value if one exists but asynchronously executes the call and updates the cache.
 * If a expireFunc is included then the cache update only happens if the cache has expired.
 *
 * @param store Use local or session storage
 * @param keyFactory: a function that returns the key for the cache value, if not provided a default hash of the url will be used
 * @param expireFunc: a function that returns a date of expiration for the cache value, if not provided the cache never expires but is always updated.
 */
function CachingPessimisticRefresh(props) {
    return (instance) => {
        const pre = async function (url, init, result) {
            const [shouldCache, getCachedValue, setCachedValue] = (0,_caching_js__WEBPACK_IMPORTED_MODULE_2__.bindCachingCore)(url, init, props);
            if (!shouldCache) {
                return [url, init, result];
            }
            const cached = getCachedValue();
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.objectDefinedNotNull)(cached)) {
                // set our result
                result = cached;
                setTimeout(async () => {
                    const q = new _queryable_js__WEBPACK_IMPORTED_MODULE_1__.Queryable(this);
                    const a = q.on.pre.toArray();
                    q.on.pre.clear();
                    // filter out this pre handler from the original queryable as we don't want to re-run it
                    a.filter(v => v !== pre).map(v => q.on.pre(v));
                    // in this case the init should contain the correct "method"
                    const value = await q(init);
                    setCachedValue(value);
                }, 0);
            }
            else {
                // register the post handler to cache the value as there is not one already in the cache
                // and we need to run this request as normal
                this.on.post((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.noInherit)(async function (url, result) {
                    setCachedValue(result);
                    return [url, result];
                }));
            }
            return [url, init, result];
        };
        instance.on.pre(pre);
        return instance;
    };
}


/***/ }),

/***/ 6936:
/*!**********************************************************!*\
  !*** ./node_modules/@pnp/queryable/behaviors/caching.js ***!
  \**********************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   CacheAlways: () => (/* binding */ CacheAlways),
/* harmony export */   CacheKey: () => (/* binding */ CacheKey),
/* harmony export */   CacheNever: () => (/* binding */ CacheNever),
/* harmony export */   bindCachingCore: () => (/* binding */ bindCachingCore)
/* harmony export */ });
/* unused harmony export Caching */
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);

/**
 * Behavior that forces caching for the request regardless of "method"
 *
 * @returns TimelinePipe
 */
function CacheAlways() {
    return (instance) => {
        instance.on.pre.prepend(async function (url, init, result) {
            init.headers = { ...init.headers, "X-PnP-CacheAlways": "1" };
            return [url, init, result];
        });
        return instance;
    };
}
/**
 * Behavior that blocks caching for the request regardless of "method"
 *
 * Note: If both Caching and CacheAlways are present AND CacheNever is present the request will not be cached
 * as we give priority to the CacheNever case
 *
 * @returns TimelinePipe
 */
function CacheNever() {
    return (instance) => {
        instance.on.pre.prepend(async function (url, init, result) {
            init.headers = { ...init.headers, "X-PnP-CacheNever": "1" };
            return [url, init, result];
        });
        return instance;
    };
}
/**
 * Behavior that allows you to specify a cache key for a request
 *
 * @param key The key to use for caching
  */
function CacheKey(key) {
    return (instance) => {
        instance.on.pre.prepend(async function (url, init, result) {
            init.headers = { ...init.headers, "X-PnP-CacheKey": key };
            return [url, init, result];
        });
        return instance;
    };
}
/**
 * Adds caching to the requests based on the supplied props
 *
 * @param props Optional props that configure how caching will work
 * @returns TimelinePipe used to configure requests
 */
function Caching(props) {
    return (instance) => {
        instance.on.pre(async function (url, init, result) {
            const [shouldCache, getCachedValue, setCachedValue] = bindCachingCore(url, init, props);
            // only cache get requested data or where the CacheAlways header is present (allows caching of POST requests)
            if (shouldCache) {
                const cached = getCachedValue();
                // we need to ensure that result stays "undefined" unless we mean to set null as the result
                if (cached === null) {
                    // if we don't have a cached result we need to get it after the request is sent. Get the raw value (un-parsed) to store into cache
                    this.on.post((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.noInherit)(async function (url, result) {
                        setCachedValue(result);
                        return [url, result];
                    }));
                }
                else {
                    result = cached;
                }
            }
            return [url, init, result];
        });
        return instance;
    };
}
const storage = new _pnp_core__WEBPACK_IMPORTED_MODULE_0__.PnPClientStorage();
/**
 * Based on the supplied properties, creates bound logic encapsulating common caching configuration
 * sharable across implementations to more easily provide consistent behavior across behaviors
 *
 * @param props Any caching props used to initialize the core functions
 */
function bindCachingCore(url, init, props) {
    var _a, _b;
    const { store, keyFactory, expireFunc } = {
        store: "local",
        keyFactory: (url) => (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.getHashCode)(url.toLowerCase()).toString(),
        expireFunc: () => (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.dateAdd)(new Date(), "minute", 5),
        ...props,
    };
    const s = store === "session" ? storage.session : storage.local;
    const key = (init === null || init === void 0 ? void 0 : init.headers["X-PnP-CacheKey"]) ? init.headers["X-PnP-CacheKey"] : keyFactory(url);
    return [
        // calculated value indicating if we should cache this request
        (/get/i.test(init.method) || ((_a = init === null || init === void 0 ? void 0 : init.headers["X-PnP-CacheAlways"]) !== null && _a !== void 0 ? _a : false)) && !((_b = init === null || init === void 0 ? void 0 : init.headers["X-PnP-CacheNever"]) !== null && _b !== void 0 ? _b : false),
        // gets the cached value
        () => s.get(key),
        // sets the cached value
        (value) => s.put(key, value, expireFunc(url)),
    ];
}


/***/ }),

/***/ 238:
/*!*************************************************************!*\
  !*** ./node_modules/@pnp/queryable/behaviors/cancelable.js ***!
  \*************************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   CancelAction: () => (/* binding */ CancelAction),
/* harmony export */   cancelableScope: () => (/* binding */ cancelableScope)
/* harmony export */ });
/* unused harmony exports asCancelableScope, Cancelable */
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);

/**
 * Cancelable is a fairly complex behavior as there is a lot to consider through multiple timelines. We have
 * two main cases:
 *
 * 1. basic method that is a single call and returns the result of an operation (return spPost(...))
 * 2. complex method that has multiple async calls within
 *
 * 1. For basic calls the cancel info is attached in init as it is only involved within a single request.
 *    This works because there is only one request and the cancel logic doesn't need to persist across
 *    inheriting instances. Also, many of these requests are so fast canceling is likely unnecessary
 *
 * 2. Complex method present a larger challenge because they are comprised of > 1 request and the promise
 *    that is actually returned to the user is not directly from one of our calls. This promise is the
 *    one "created" by the language when you await. For complex methods we have two things that solve these
 *    needs.
 *
 *    The first is the use of either the cancelableScope decorator or the asCancelableScope method
 *    wrapper. These create an upper level cancel info that is then shared across the child requests within
 *    the complex method. Meaning if I do a files.addChunked the same cancel info (and cancel method)
 *    are set on the current "this" which is user object on which the method was called. This info is then
 *    passed down to any child requests using the original "this" as a base using the construct moment.
 *
 *    The CancelAction behavior is used to apply additional actions to a request once it is canceled. For example
 *    in the case of uploading files chunked in sp we cancel the upload by id.
 */
// this is a special moment used to broadcast when a request is canceled
const MomentName = "__CancelMoment__";
// this value is used to track cancel state and the value is represetented by IScopeInfo
const ScopeId = Symbol.for("CancelScopeId");
// module map of all currently tracked cancel scopes
const cancelScopes = new Map();
/**
 * This method is bound to a scope id and used as the cancel method exposed to the user via cancelable promise
 *
 * @param this unused, the current promise
 * @param scopeId Id bound at creation time
 */
async function cancelPrimitive(scopeId) {
    const scope = cancelScopes.get(scopeId);
    scope.controller.abort();
    if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isArray)(scope === null || scope === void 0 ? void 0 : scope.actions)) {
        scope.actions.map(action => scope.currentSelf.on[MomentName](action));
    }
    try {
        await scope.currentSelf.emit[MomentName]();
    }
    catch (e) {
        scope.currentSelf.log(`Error in cancel: ${e}`, 3);
    }
}
/**
 * Creates a new scope id, sets it on the instance's ScopeId property, and adds the info to the map
 *
 * @returns the new scope id (GUID)
 */
function createScope(instance) {
    const id = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.getGUID)();
    instance[ScopeId] = id;
    cancelScopes.set(id, {
        cancel: cancelPrimitive.bind({}, id),
        actions: [],
        controller: null,
        currentSelf: instance,
    });
    return id;
}
/**
 * Function wrapper that turns the supplied function into a cancellation scope
 *
 * @param func Func to wrap
 * @returns The same func signature, wrapped with our cancel scoping logic
 */
const asCancelableScope = (func) => {
    return function (...args) {
        // ensure we have setup "this" to cancel
        // 1. for single requests the value is set in the behavior's init observer
        // 2. for complex requests the value is set here
        if (!Reflect.has(this, ScopeId)) {
            createScope(this);
        }
        // execute the original function, but don't await it
        const result = func.apply(this, args).finally(() => {
            // remove any cancel scope values tied to this instance
            cancelScopes.delete(this[ScopeId]);
            delete this[ScopeId];
        });
        // ensure the synthetic promise from a complex method has a cancel method
        result.cancel = cancelScopes.get(this[ScopeId]).cancel;
        return result;
    };
};
/**
 * Decorator used to mark multi-step methods to ensure all subrequests are properly cancelled
 */
function cancelableScope(_target, _propertyKey, descriptor) {
    // wrapping the original method
    descriptor.value = asCancelableScope(descriptor.value);
}
/**
 * Allows requests to be canceled by the caller by adding a cancel method to the Promise returned by the library
 *
 * @returns Timeline pipe to setup canelability
 */
function Cancelable() {
    if (!AbortController) {
        throw Error("The current environment appears to not support AbortController, please include a suitable polyfill.");
    }
    return (instance) => {
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        instance.on.construct(function (init, path) {
            if (typeof init !== "string") {
                const parent = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isArray)(init) ? init[0] : init;
                if (Reflect.has(parent, ScopeId)) {
                    // ensure we carry over the scope id to the new instance from the parent
                    this[ScopeId] = parent[ScopeId];
                }
                // define the moment's implementation
                this.moments[MomentName] = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.asyncBroadcast)();
            }
        });
        // init our queryable to support cancellation
        instance.on.init(function () {
            if (!Reflect.has(this, ScopeId)) {
                // ensure we have setup "this" to cancel
                // 1. for single requests this will set the value
                // 2. for complex requests the value is set in asCancelableScope
                const id = createScope(this);
                // if we are creating the scope here, we have not created it within asCancelableScope
                // meaning the finally handler there will not delete the tracked scope reference
                this.on.dispose(() => {
                    cancelScopes.delete(id);
                });
            }
            this.on[this.InternalPromise]((promise) => {
                // when a new promise is created add a cancel method
                promise.cancel = cancelScopes.get(this[ScopeId]).cancel;
                return [promise];
            });
        });
        instance.on.pre(async function (url, init, result) {
            // grab the current scope, update the controller and currentSelf
            const existingScope = cancelScopes.get(this[ScopeId]);
            // if we are here without a scope we are likely running a CancelAction request so we just ignore canceling
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.objectDefinedNotNull)(existingScope)) {
                const controller = new AbortController();
                existingScope.controller = controller;
                existingScope.currentSelf = this;
                if (init.signal) {
                    // we do our best to hook our logic to the existing signal
                    init.signal.addEventListener("abort", () => {
                        existingScope.cancel();
                    });
                }
                else {
                    init.signal = controller.signal;
                }
            }
            return [url, init, result];
        });
        // clean up any cancel info from the object after the request lifecycle is complete
        instance.on.dispose(function () {
            delete this[ScopeId];
            delete this.moments[MomentName];
        });
        return instance;
    };
}
/**
 * Allows you to define an action that is run when a request is cancelled
 *
 * @param action The action to run
 * @returns A timeline pipe used in the request lifecycle
 */
function CancelAction(action) {
    return (instance) => {
        instance.on.pre(async function (...args) {
            const existingScope = cancelScopes.get(this[ScopeId]);
            // if we don't have a scope this request is not using Cancelable so we do nothing
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.objectDefinedNotNull)(existingScope)) {
                if (!(0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isArray)(existingScope.actions)) {
                    existingScope.actions = [];
                }
                if (existingScope.actions.indexOf(action) < 0) {
                    existingScope.actions.push(action);
                }
            }
            return args;
        });
        return instance;
    };
}


/***/ }),

/***/ 6046:
/*!*****************************************************************!*\
  !*** ./node_modules/@pnp/queryable/behaviors/inject-headers.js ***!
  \*****************************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   InjectHeaders: () => (/* binding */ InjectHeaders)
/* harmony export */ });
function InjectHeaders(headers, prepend = false) {
    return (instance) => {
        const f = async function (url, init, result) {
            init.headers = { ...init.headers, ...headers };
            return [url, init, result];
        };
        if (prepend) {
            instance.on.pre.prepend(f);
        }
        else {
            instance.on.pre(f);
        }
        return instance;
    };
}


/***/ }),

/***/ 5102:
/*!**********************************************************!*\
  !*** ./node_modules/@pnp/queryable/behaviors/parsers.js ***!
  \**********************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   BlobParse: () => (/* binding */ BlobParse),
/* harmony export */   BufferParse: () => (/* binding */ BufferParse),
/* harmony export */   DefaultParse: () => (/* binding */ DefaultParse),
/* harmony export */   HttpRequestError: () => (/* binding */ HttpRequestError),
/* harmony export */   JSONParse: () => (/* binding */ JSONParse),
/* harmony export */   TextParse: () => (/* binding */ TextParse),
/* harmony export */   parseBinderWithErrorCheck: () => (/* binding */ parseBinderWithErrorCheck),
/* harmony export */   parseODataJSON: () => (/* binding */ parseODataJSON)
/* harmony export */ });
/* unused harmony exports HeaderParse, JSONHeaderParse, errorCheck */
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);


function DefaultParse() {
    return parseBinderWithErrorCheck(async (response) => {
        // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
        if ((response.headers.has("Content-Length") && parseFloat(response.headers.get("Content-Length")) === 0) || response.status === 204) {
            return {};
        }
        // patch to handle cases of 200 response with no or whitespace only bodies (#487 & #545)
        const txt = await response.text();
        const json = txt.replace(/\s/ig, "").length > 0 ? JSON.parse(txt) : {};
        return parseODataJSON(json);
    });
}
function TextParse() {
    return parseBinderWithErrorCheck(r => r.text());
}
function BlobParse() {
    return parseBinderWithErrorCheck(r => r.blob());
}
function JSONParse() {
    return parseBinderWithErrorCheck(r => r.json());
}
function BufferParse() {
    return parseBinderWithErrorCheck(r => (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isFunc)(r.arrayBuffer) ? r.arrayBuffer() : r.buffer());
}
function HeaderParse() {
    return parseBinderWithErrorCheck(async (r) => r.headers);
}
function JSONHeaderParse() {
    return parseBinderWithErrorCheck(async (response) => {
        // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
        if ((response.headers.has("Content-Length") && parseFloat(response.headers.get("Content-Length")) === 0) || response.status === 204) {
            return {};
        }
        // patch to handle cases of 200 response with no or whitespace only bodies (#487 & #545)
        const txt = await response.text();
        const json = txt.replace(/\s/ig, "").length > 0 ? JSON.parse(txt) : {};
        return { data: { ...parseODataJSON(json) }, headers: { ...response.headers } };
    });
}
async function errorCheck(url, response, result) {
    if (!response.ok) {
        throw await HttpRequestError.init(response);
    }
    return [url, response, result];
}
function parseODataJSON(json) {
    let result = json;
    if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(json, "d")) {
        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(json.d, "results")) {
            result = json.d.results;
        }
        else {
            result = json.d;
        }
    }
    else if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(json, "value")) {
        result = json.value;
    }
    return result;
}
/**
 * Provides a clean way to create new parse bindings without having to duplicate a lot of boilerplate
 * Includes errorCheck ahead of the supplied impl
 *
 * @param impl Method used to parse the response
 * @returns Queryable behavior binding function
 */
function parseBinderWithErrorCheck(impl) {
    return (instance) => {
        // we clear anything else registered for parse
        // add error check
        // add the impl function we are supplied
        instance.on.parse.replace(errorCheck);
        instance.on.parse(async (url, response, result) => {
            if (response.ok && typeof result === "undefined") {
                result = await impl(response);
            }
            return [url, response, result];
        });
        return instance;
    };
}
class HttpRequestError extends Error {
    constructor(message, response, status = response.status, statusText = response.statusText) {
        super(message);
        this.response = response;
        this.status = status;
        this.statusText = statusText;
        this.isHttpRequestError = true;
    }
    static async init(r) {
        const t = await r.clone().text();
        return new HttpRequestError(`Error making HttpClient request in queryable [${r.status}] ${r.statusText} ::> ${t}`, r);
    }
}


/***/ }),

/***/ 5234:
/*!************************************************************!*\
  !*** ./node_modules/@pnp/queryable/behaviors/resolvers.js ***!
  \************************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   RejectOnError: () => (/* binding */ RejectOnError),
/* harmony export */   ResolveOnData: () => (/* binding */ ResolveOnData)
/* harmony export */ });
function ResolveOnData() {
    return (instance) => {
        instance.on.data(function (data) {
            this.emit[this.InternalResolve](data);
        });
        return instance;
    };
}
function RejectOnError() {
    return (instance) => {
        instance.on.error(function (err) {
            this.emit[this.InternalReject](err);
        });
        return instance;
    };
}


/***/ }),

/***/ 6844:
/*!**********************************************!*\
  !*** ./node_modules/@pnp/queryable/index.js ***!
  \**********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   BlobParse: () => (/* reexport safe */ _behaviors_parsers_js__WEBPACK_IMPORTED_MODULE_7__.BlobParse),
/* harmony export */   BrowserFetchWithRetry: () => (/* reexport safe */ _behaviors_browser_fetch_js__WEBPACK_IMPORTED_MODULE_2__.BrowserFetchWithRetry),
/* harmony export */   BufferParse: () => (/* reexport safe */ _behaviors_parsers_js__WEBPACK_IMPORTED_MODULE_7__.BufferParse),
/* harmony export */   CacheAlways: () => (/* reexport safe */ _behaviors_caching_js__WEBPACK_IMPORTED_MODULE_3__.CacheAlways),
/* harmony export */   CacheKey: () => (/* reexport safe */ _behaviors_caching_js__WEBPACK_IMPORTED_MODULE_3__.CacheKey),
/* harmony export */   CacheNever: () => (/* reexport safe */ _behaviors_caching_js__WEBPACK_IMPORTED_MODULE_3__.CacheNever),
/* harmony export */   CancelAction: () => (/* reexport safe */ _behaviors_cancelable_js__WEBPACK_IMPORTED_MODULE_5__.CancelAction),
/* harmony export */   DefaultParse: () => (/* reexport safe */ _behaviors_parsers_js__WEBPACK_IMPORTED_MODULE_7__.DefaultParse),
/* harmony export */   InjectHeaders: () => (/* reexport safe */ _behaviors_inject_headers_js__WEBPACK_IMPORTED_MODULE_6__.InjectHeaders),
/* harmony export */   JSONParse: () => (/* reexport safe */ _behaviors_parsers_js__WEBPACK_IMPORTED_MODULE_7__.JSONParse),
/* harmony export */   Queryable: () => (/* reexport safe */ _queryable_js__WEBPACK_IMPORTED_MODULE_1__.Queryable),
/* harmony export */   RejectOnError: () => (/* reexport safe */ _behaviors_resolvers_js__WEBPACK_IMPORTED_MODULE_8__.RejectOnError),
/* harmony export */   ResolveOnData: () => (/* reexport safe */ _behaviors_resolvers_js__WEBPACK_IMPORTED_MODULE_8__.ResolveOnData),
/* harmony export */   TextParse: () => (/* reexport safe */ _behaviors_parsers_js__WEBPACK_IMPORTED_MODULE_7__.TextParse),
/* harmony export */   addProp: () => (/* binding */ addProp),
/* harmony export */   body: () => (/* binding */ body),
/* harmony export */   cancelableScope: () => (/* reexport safe */ _behaviors_cancelable_js__WEBPACK_IMPORTED_MODULE_5__.cancelableScope),
/* harmony export */   del: () => (/* reexport safe */ _queryable_js__WEBPACK_IMPORTED_MODULE_1__.del),
/* harmony export */   get: () => (/* reexport safe */ _queryable_js__WEBPACK_IMPORTED_MODULE_1__.get),
/* harmony export */   headers: () => (/* binding */ headers),
/* harmony export */   invokable: () => (/* reexport safe */ _queryable_js__WEBPACK_IMPORTED_MODULE_1__.invokable),
/* harmony export */   op: () => (/* reexport safe */ _queryable_js__WEBPACK_IMPORTED_MODULE_1__.op),
/* harmony export */   parseBinderWithErrorCheck: () => (/* reexport safe */ _behaviors_parsers_js__WEBPACK_IMPORTED_MODULE_7__.parseBinderWithErrorCheck),
/* harmony export */   parseODataJSON: () => (/* reexport safe */ _behaviors_parsers_js__WEBPACK_IMPORTED_MODULE_7__.parseODataJSON),
/* harmony export */   patch: () => (/* reexport safe */ _queryable_js__WEBPACK_IMPORTED_MODULE_1__.patch),
/* harmony export */   post: () => (/* reexport safe */ _queryable_js__WEBPACK_IMPORTED_MODULE_1__.post),
/* harmony export */   queryableFactory: () => (/* reexport safe */ _queryable_js__WEBPACK_IMPORTED_MODULE_1__.queryableFactory)
/* harmony export */ });
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _queryable_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./queryable.js */ 7111);
/* harmony import */ var _behaviors_browser_fetch_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./behaviors/browser-fetch.js */ 8786);
/* harmony import */ var _behaviors_caching_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./behaviors/caching.js */ 6936);
/* harmony import */ var _behaviors_caching_pessimistic_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./behaviors/caching-pessimistic.js */ 7395);
/* harmony import */ var _behaviors_cancelable_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./behaviors/cancelable.js */ 238);
/* harmony import */ var _behaviors_inject_headers_js__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./behaviors/inject-headers.js */ 6046);
/* harmony import */ var _behaviors_parsers_js__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ./behaviors/parsers.js */ 5102);
/* harmony import */ var _behaviors_resolvers_js__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ./behaviors/resolvers.js */ 5234);


/**
 * Behavior exports
 */










/**
 * Adds a property to a target instance
 *
 * @param target The object to whose prototype we will add a property
 * @param name Property name
 * @param factory Factory method used to produce the property value
 * @param path Any additional path required to produce the value
 */
function addProp(target, name, factory, path) {
    Reflect.defineProperty(target.prototype, name, {
        configurable: true,
        enumerable: true,
        get: function () {
            return factory(this, path || name);
        },
    });
}
/**
 * takes the supplied object of type U, JSON.stringify's it, and sets it as the value of a "body" property
 */
function body(o, previous) {
    return Object.assign({ body: (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.jsS)(o) }, previous);
}
/**
 * Adds headers to an new/existing RequestInit
 *
 * @param o Headers to add
 * @param previous Any previous partial RequestInit
 * @returns RequestInit combining previous and specified headers
 */
// eslint-disable-next-line @typescript-eslint/ban-types
function headers(o, previous) {
    return Object.assign({}, previous, { headers: { ...previous === null || previous === void 0 ? void 0 : previous.headers, ...o } });
}


/***/ }),

/***/ 7111:
/*!**************************************************!*\
  !*** ./node_modules/@pnp/queryable/queryable.js ***!
  \**************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Queryable: () => (/* binding */ Queryable),
/* harmony export */   del: () => (/* binding */ del),
/* harmony export */   get: () => (/* binding */ get),
/* harmony export */   invokable: () => (/* binding */ invokable),
/* harmony export */   op: () => (/* binding */ op),
/* harmony export */   patch: () => (/* binding */ patch),
/* harmony export */   post: () => (/* binding */ post),
/* harmony export */   queryableFactory: () => (/* binding */ queryableFactory)
/* harmony export */ });
/* unused harmony export put */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! tslib */ 4340);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);


const DefaultMoments = {
    construct: (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.lifecycle)(),
    pre: (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.asyncReduce)(),
    auth: (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.asyncReduce)(),
    send: (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.request)(),
    parse: (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.asyncReduce)(),
    post: (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.asyncReduce)(),
    data: (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.broadcast)(),
};
let Queryable = class Queryable extends _pnp_core__WEBPACK_IMPORTED_MODULE_0__.Timeline {
    constructor(init, path) {
        super(DefaultMoments);
        // these keys represent internal events for Queryable, users are not expected to
        // subscribe directly to these, rather they enable functionality within Queryable
        // they are Symbols such that there are NOT cloned between queryables as we only grab string keys (by design)
        this.InternalResolve = Symbol.for("Queryable_Resolve");
        this.InternalReject = Symbol.for("Queryable_Reject");
        this.InternalPromise = Symbol.for("Queryable_Promise");
        // default to use the included URL search params to parse the query string
        this._query = new URLSearchParams();
        // add an internal moment with specific implementation for promise creation
        this.moments[this.InternalPromise] = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.reduce)();
        let parent;
        if (typeof init === "string") {
            this._url = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.combine)(init, path);
        }
        else if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isArray)(init)) {
            if (init.length !== 2) {
                throw Error("When using the tuple param exactly two arguments are expected.");
            }
            if (typeof init[1] !== "string") {
                throw Error("Expected second tuple param to be a string.");
            }
            parent = init[0];
            this._url = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.combine)(init[1], path);
        }
        else {
            parent = init;
            this._url = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.combine)(parent._url, path);
        }
        if (typeof parent !== "undefined") {
            this.observers = parent.observers;
            this._inheritingObservers = true;
        }
    }
    /**
     * Directly concatenates the supplied string to the current url, not normalizing "/" chars
     *
     * @param pathPart The string to concatenate to the url
     */
    concat(pathPart) {
        this._url += pathPart;
        return this;
    }
    /**
     * Gets the full url with query information
     *
     */
    toRequestUrl() {
        let url = this.toUrl();
        const query = this.query.toString();
        if (!(0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.stringIsNullOrEmpty)(query)) {
            url += `${url.indexOf("?") > -1 ? "&" : "?"}${query}`;
        }
        return url;
    }
    /**
     * Querystring key, value pairs which will be included in the request
     */
    get query() {
        return this._query;
    }
    /**
     * Gets the current url
     *
     */
    toUrl() {
        return this._url;
    }
    execute(userInit) {
        // if there are NO observers registered this is likely either a bug in the library or a user error, direct to docs
        if (Reflect.ownKeys(this.observers).length < 1) {
            throw Error("No observers registered for this request. (https://pnp.github.io/pnpjs/queryable/queryable#no-observers-registered-for-this-request)");
        }
        // schedule the execution after we return the promise below in the next event loop
        setTimeout(async () => {
            const requestId = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.getGUID)();
            let requestUrl;
            const log = (msg, level) => {
                // this allows us to easily and consistently format our messages
                this.log(`[${requestId}] ${msg}`, level);
            };
            try {
                log("Beginning request", 0);
                // include the request id in the headers to assist with debugging against logs
                const initSeed = {
                    ...userInit,
                    headers: { ...userInit.headers, "X-PnPjs-RequestId": requestId },
                };
                // eslint-disable-next-line prefer-const
                let [url, init, result] = await this.emit.pre(this.toRequestUrl(), initSeed, undefined);
                log(`Url: ${url}`, 1);
                if (typeof result !== "undefined") {
                    log("Result returned from pre, Emitting data");
                    this.emit.data(result);
                    log("Emitted data");
                    return;
                }
                log("Emitting auth");
                [requestUrl, init] = await this.emit.auth(new URL(url), init);
                log("Emitted auth");
                // we always resepect user supplied init over observer modified init
                init = { ...init, ...userInit, headers: { ...init.headers, ...userInit.headers } };
                log("Emitting send");
                let response = await this.emit.send(requestUrl, init);
                log("Emitted send");
                log("Emitting parse");
                [requestUrl, response, result] = await this.emit.parse(requestUrl, response, result);
                log("Emitted parse");
                log("Emitting post");
                [requestUrl, result] = await this.emit.post(requestUrl, result);
                log("Emitted post");
                log("Emitting data");
                this.emit.data(result);
                log("Emitted data");
            }
            catch (e) {
                log(`Emitting error: "${e.message || e}"`, 3);
                // anything that throws we emit and continue
                this.error(e);
                log("Emitted error", 3);
            }
            finally {
                log("Finished request", 0);
            }
        }, 0);
        // this allows us to internally hook the promise creation and modify it. This was introduced to allow for
        // cancelable to work as envisioned, but may have other users. Meant for internal use in the library accessed via behaviors.
        return this.emit[this.InternalPromise](new Promise((resolve, reject) => {
            // we overwrite any pre-existing internal events as a
            // given queryable only processes a single request at a time
            this.on[this.InternalResolve].replace(resolve);
            this.on[this.InternalReject].replace(reject);
        }))[0];
    }
};
Queryable = (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__decorate)([
    invokable()
    // eslint-disable-next-line @typescript-eslint/no-unsafe-declaration-merging
], Queryable);

function ensureInit(method, init = { headers: {} }) {
    return { method, ...init, headers: { ...init.headers } };
}
function get(init) {
    return this.start(ensureInit("GET", init));
}
function post(init) {
    return this.start(ensureInit("POST", init));
}
function put(init) {
    return this.start(ensureInit("PUT", init));
}
function patch(init) {
    return this.start(ensureInit("PATCH", init));
}
function del(init) {
    return this.start(ensureInit("DELETE", init));
}
function op(q, operation, init) {
    return Reflect.apply(operation, q, [init]);
}
function queryableFactory(constructor) {
    return (init, path) => {
        // construct the concrete instance
        const instance = new constructor(init, path);
        // we emit the construct event from the factory because we need all of the decorators and constructors
        // to have fully finished before we emit, which is now true. We type the instance to any to get around
        // the protected nature of emit
        instance.emit.construct(init, path);
        return instance;
    };
}
/**
 * Allows a decorated object to be invoked as a function, optionally providing an implementation for that action
 *
 * @param invokeableAction Optional. The logic to execute upon invoking the object as a function.
 * @returns Decorator which applies the invokable logic to the tagged class
 */
function invokable(invokeableAction) {
    return (target) => {
        return new Proxy(target, {
            construct(clz, args, newTarget) {
                const invokableInstance = Object.assign(function (init) {
                    if (!(0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isFunc)(invokeableAction)) {
                        invokeableAction = function (init) {
                            return op(this, get, init);
                        };
                    }
                    return Reflect.apply(invokeableAction, invokableInstance, [init]);
                }, Reflect.construct(clz, args, newTarget));
                Reflect.setPrototypeOf(invokableInstance, newTarget.prototype);
                return invokableInstance;
            },
        });
    };
}


/***/ }),

/***/ 7858:
/*!**************************************************!*\
  !*** ./node_modules/@pnp/sp/appcatalog/index.js ***!
  \**************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _fi_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../fi.js */ 6553);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 2315);
/* harmony import */ var _web_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./web.js */ 9428);





Reflect.defineProperty(_fi_js__WEBPACK_IMPORTED_MODULE_0__.SPFI.prototype, "tenantAppcatalog", {
    configurable: true,
    enumerable: true,
    get: function () {
        return this.create(_types_js__WEBPACK_IMPORTED_MODULE_2__.AppCatalog, "_api/web/tenantappcatalog/AvailableApps");
    },
});
_fi_js__WEBPACK_IMPORTED_MODULE_0__.SPFI.prototype.getTenantAppCatalogWeb = async function () {
    const data = await (0,_webs_types_js__WEBPACK_IMPORTED_MODULE_1__.Web)(this._root, "_api/SP_TenantSettings_Current")();
    return (0,_webs_types_js__WEBPACK_IMPORTED_MODULE_1__.Web)([this._root, data.CorporateCatalogUrl]);
};


/***/ }),

/***/ 2315:
/*!**************************************************!*\
  !*** ./node_modules/@pnp/sp/appcatalog/types.js ***!
  \**************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   AppCatalog: () => (/* binding */ AppCatalog)
/* harmony export */ });
/* unused harmony exports _AppCatalog, _App, App */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../utils/extract-web-url.js */ 2997);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../decorators.js */ 6540);





function getAppCatalogPath(base, path) {
    const paths = ["_api/web/tenantappcatalog/", "_api/web/sitecollectionappcatalog/"];
    for (let i = 0; i < paths.length; i++) {
        const index = base.indexOf(paths[i]);
        if (index > -1) {
            return (0,_pnp_core__WEBPACK_IMPORTED_MODULE_2__.combine)(base.substring(index, index + paths[i].length), path);
        }
    }
    return (0,_pnp_core__WEBPACK_IMPORTED_MODULE_2__.combine)(base, path);
}
let _AppCatalog = class _AppCatalog extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPCollection {
    constructor(base, path) {
        super(base, null);
        this._url = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_2__.combine)((0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_1__.extractWebUrl)(this._url), path);
    }
    /**
     * Get details of specific app from the app catalog
     * @param id - Specify the guid of the app
     */
    getAppById(id) {
        return App(this, `getById('${id}')`);
    }
    /**
     * Synchronize a solution to the Microsoft Teams App Catalog
     * @param id - Specify the guid of the app
     * @param useSharePointItemId (optional) - By default this REST call requires the SP Item id of the app, not the app id.
     *                            PnPjs will try to fetch the item id, you can still use this parameter to pass your own item id in the first parameter
     */
    async syncSolutionToTeams(id, useSharePointItemId = false) {
        // This REST call requires that you refer the list item id of the solution in the app catalog site.
        let appId = null;
        const webUrl = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_2__.combine)((0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_1__.extractWebUrl)(this.toUrl()), "_api/web");
        if (useSharePointItemId) {
            appId = id;
        }
        else {
            const listId = (await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.SPCollection)([this, webUrl], "lists").select("Id").filter("EntityTypeName eq 'AppCatalog'")())[0].Id;
            const listItems = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.SPCollection)([this, webUrl], `lists/getById('${listId}')/items`).select("Id").filter(`AppProductID eq '${id}'`).top(1)();
            if (listItems && listItems.length > 0) {
                appId = listItems[0].Id;
            }
            else {
                throw Error(`Did not find the app with id ${id} in the appcatalog.`);
            }
        }
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(AppCatalog(this, getAppCatalogPath(this.toUrl(), `SyncSolutionToTeams(id=${appId})`)));
    }
    /**
     * Uploads an app package. Not supported for batching
     *
     * @param filename Filename to create.
     * @param content app package data (eg: the .app or .sppkg file).
     * @param shouldOverWrite Should an app with the same name in the same location be overwritten? (default: true)
     * @returns Promise<IAppAddResult>
     */
    async add(filename, content, shouldOverWrite = true) {
        // you don't add to the availableapps collection
        const adder = AppCatalog(this, getAppCatalogPath(this.toUrl(), `add(overwrite=${shouldOverWrite},url='${filename}')`));
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(adder, {
            body: content, headers: {
                "binaryStringRequestBody": "true",
            },
        });
    }
};
_AppCatalog = (0,tslib__WEBPACK_IMPORTED_MODULE_3__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_4__.defaultPath)("_api/web/tenantappcatalog/AvailableApps")
], _AppCatalog);

const AppCatalog = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_AppCatalog);
class _App extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPInstance {
    /**
     * This method deploys an app on the app catalog. It must be called in the context
     * of the tenant app catalog web or it will fail.
     *
     * @param skipFeatureDeployment Deploy the app to the entire tenant
     */
    deploy(skipFeatureDeployment = false) {
        return this.do(`Deploy(${skipFeatureDeployment})`);
    }
    /**
     * This method retracts a deployed app on the app catalog. It must be called in the context
     * of the tenant app catalog web or it will fail.
     */
    retract() {
        return this.do("Retract");
    }
    /**
     * This method allows an app which is already deployed to be installed on a web
     */
    install() {
        return this.do("Install");
    }
    /**
     * This method allows an app which is already installed to be uninstalled on a web
     * Note: when you use the REST API to uninstall a solution package from the site, it is not relocated to the recycle bin
     */
    uninstall() {
        return this.do("Uninstall");
    }
    /**
     * This method allows an app which is already installed to be upgraded on a web
     */
    upgrade() {
        return this.do("Upgrade");
    }
    /**
     * This method removes an app from the app catalog. It must be called in the context
     * of the tenant app catalog web or it will fail.
     */
    remove() {
        return this.do("Remove");
    }
    do(path) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(App(this, path));
    }
}
const App = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_App);


/***/ }),

/***/ 9428:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/appcatalog/web.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 2315);



// we use this function to wrap the AppCatalog as we want to ignore any path values addProp
// will pass and use the defaultPath defined for AppCatalog
(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web, "appcatalog", (s) => (0,_types_js__WEBPACK_IMPORTED_MODULE_2__.AppCatalog)(s, "_api/web/sitecollectionappcatalog/AvailableApps"));


/***/ }),

/***/ 2235:
/*!***************************************************!*\
  !*** ./node_modules/@pnp/sp/attachments/index.js ***!
  \***************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _item_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./item.js */ 8833);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 5050);




/***/ }),

/***/ 8833:
/*!**************************************************!*\
  !*** ./node_modules/@pnp/sp/attachments/item.js ***!
  \**************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _items_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../items/types.js */ 132);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 5050);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_items_types_js__WEBPACK_IMPORTED_MODULE_1__._Item, "attachmentFiles", _types_js__WEBPACK_IMPORTED_MODULE_2__.Attachments);


/***/ }),

/***/ 5050:
/*!***************************************************!*\
  !*** ./node_modules/@pnp/sp/attachments/types.js ***!
  \***************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Attachments: () => (/* binding */ Attachments)
/* harmony export */ });
/* unused harmony exports _Attachments, _Attachment, Attachment */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../decorators.js */ 6540);
/* harmony import */ var _files_readable_file_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../files/readable-file.js */ 3645);
/* harmony import */ var _utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../utils/encode-path-str.js */ 4729);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../spqueryable.js */ 2678);






let _Attachments = class _Attachments extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_3__._SPCollection {
    /**
    * Gets a Attachment File by filename
    *
    * @param name The name of the file, including extension.
    */
    getByName(name) {
        const f = Attachment(this);
        f.concat(`('${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_2__.encodePath)(name)}')`);
        return f;
    }
    /**
     * Adds a new attachment to the collection. Not supported for batching.
     *
     * @param name The name of the file, including extension.
     * @param content The Base64 file content.
     */
    async add(name, content) {
        const response = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.spPost)(Attachments(this, `add(FileName='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_2__.encodePath)(name)}')`), { body: content });
        return {
            data: response,
            file: this.getByName(name),
        };
    }
};
_Attachments = (0,tslib__WEBPACK_IMPORTED_MODULE_4__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_5__.defaultPath)("AttachmentFiles")
], _Attachments);

const Attachments = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.spInvokableFactory)(_Attachments);
class _Attachment extends _files_readable_file_js__WEBPACK_IMPORTED_MODULE_1__.ReadableFile {
    constructor() {
        super(...arguments);
        this.delete = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.deleteableWithETag)();
    }
    /**
     * Sets the content of a file. Not supported for batching
     *
     * @param content The value to set for the file contents
     */
    async setContent(body) {
        await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.spPost)(Attachment(this, "$value"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.headers)({ "X-HTTP-Method": "PUT" }, { body }));
        return this;
    }
    /**
     * Delete this attachment file and send it to recycle bin
     *
     * @param eTag Value used in the IF-Match header, by default "*"
     */
    recycle(eTag = "*") {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.spPost)(Attachment(this, "recycleObject"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.headers)({
            "IF-Match": eTag,
            "X-HTTP-Method": "DELETE",
        }));
    }
}
const Attachment = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.spInvokableFactory)(_Attachment);


/***/ }),

/***/ 8018:
/*!******************************************!*\
  !*** ./node_modules/@pnp/sp/batching.js ***!
  \******************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   BatchNever: () => (/* binding */ BatchNever),
/* harmony export */   createBatch: () => (/* binding */ createBatch)
/* harmony export */ });
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./spqueryable.js */ 2678);
/* harmony import */ var _fi_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./fi.js */ 6553);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./webs/types.js */ 3169);





_fi_js__WEBPACK_IMPORTED_MODULE_3__.SPFI.prototype.batched = function (props) {
    const batched = (0,_fi_js__WEBPACK_IMPORTED_MODULE_3__.spfi)(this);
    const [behavior, execute] = createBatch(batched._root, props);
    batched.using(behavior);
    return [batched, execute];
};
_webs_types_js__WEBPACK_IMPORTED_MODULE_4__._Web.prototype.batched = function (props) {
    const batched = (0,_webs_types_js__WEBPACK_IMPORTED_MODULE_4__.Web)(this);
    const [behavior, execute] = createBatch(batched, props);
    batched.using(behavior);
    return [batched, execute];
};
/**
 * Tracks on a batched instance that registration is complete (the child request has gotten to the send moment and the request is included in the batch)
 */
const RegistrationCompleteSym = Symbol.for("batch_registration");
/**
 * Tracks on a batched instance that the child request timeline lifecycle is complete (called in child.dispose)
 */
const RequestCompleteSym = Symbol.for("batch_request");
/**
 * Special batch parsing behavior used to convert the batch response text into a set of Response objects for each request
 * @returns A parser behavior
 */
function BatchParse() {
    return (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.parseBinderWithErrorCheck)(async (response) => {
        const text = await response.text();
        return parseResponse(text);
    });
}
/**
 * Internal class used to execute the batch request through the timeline lifecycle
 */
class BatchQueryable extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_2__._SPQueryable {
    constructor(base, requestBaseUrl = base.toUrl().replace(/_api[\\|/].*$/i, "")) {
        super(requestBaseUrl, "_api/$batch");
        this.requestBaseUrl = requestBaseUrl;
        // this will copy over the current observables from the base associated with this batch
        // this will replace any other parsing present
        this.using((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.CopyFrom)(base, "replace"), BatchParse());
        this.on.dispose(() => {
            // there is a code path where you may invoke a batch, say on items.add, whose return
            // is an object like { data: any, item: IItem }. The expectation from v1 on is `item` in that object
            // is immediately usable to make additional queries. Without this step when that IItem instance is
            // created using "this.getById" within IITems.add all of the current observers of "this" are
            // linked to the IItem instance created (expected), BUT they will be the set of observers setup
            // to handle the batch, meaning invoking `item` will result in a half batched call that
            // doesn't really work. To deliver the expected functionality we "reset" the
            // observers using the original instance, mimicing the behavior had
            // the IItem been created from that base without a batch involved. We use CopyFrom to ensure
            // that we maintain the references to the InternalResolve and InternalReject events through
            // the end of this timeline lifecycle. This works because CopyFrom by design uses Object.keys
            // which ignores symbol properties.
            base.using((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.CopyFrom)(this, "replace", (k) => /(auth|send|pre|init)/i.test(k)));
        });
    }
}
/**
 * Creates a batched version of the supplied base, meaning that all chained fluent operations from the new base are part of the batch
 *
 * @param base The base from which to initialize the batch
 * @param props Any properties used to initialize the batch functionality
 * @returns A tuple of [behavior used to assign objects to the batch, the execute function used to resolve the batch requests]
 */
function createBatch(base, props) {
    const registrationPromises = [];
    const completePromises = [];
    const requests = [];
    const batchQuery = new BatchQueryable(base);
    // this id will be reused across multiple batches if the number of requests added to the batch
    // exceeds the configured maxRequests value
    const batchId = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.getGUID)();
    // this query is used to copy back the behaviors after the batch executes
    // it should not manipulated or have behaviors added.
    const refQuery = new BatchQueryable(base);
    const { headersCopyPattern, maxRequests } = {
        headersCopyPattern: /Accept|Content-Type|IF-Match/i,
        maxRequests: 20,
        ...props,
    };
    const execute = async () => {
        await Promise.all(registrationPromises);
        if (requests.length < 1) {
            // even if we have no requests we need to await the complete promises to ensure
            // that execute only resolves AFTER every child request disposes #2457
            // this likely means caching is being used, we returned values for all child requests from the cache
            return Promise.all(completePromises).then(() => void (0));
        }
        // create a working copy of our requests
        const requestsWorkingCopy = requests.slice();
        while (requestsWorkingCopy.length > 0) {
            const requestsChunk = requestsWorkingCopy.splice(0, maxRequests);
            const batchBody = [];
            let currentChangeSetId = "";
            for (let i = 0; i < requestsChunk.length; i++) {
                const [, url, init] = requestsChunk[i];
                if (init.method === "GET") {
                    if (currentChangeSetId.length > 0) {
                        // end an existing change set
                        batchBody.push(`--changeset_${currentChangeSetId}--\n\n`);
                        currentChangeSetId = "";
                    }
                    batchBody.push(`--batch_${batchId}\n`);
                }
                else {
                    if (currentChangeSetId.length < 1) {
                        // start new change set
                        currentChangeSetId = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.getGUID)();
                        batchBody.push(`--batch_${batchId}\n`);
                        batchBody.push(`Content-Type: multipart/mixed; boundary="changeset_${currentChangeSetId}"\n\n`);
                    }
                    batchBody.push(`--changeset_${currentChangeSetId}\n`);
                }
                // common batch part prefix
                batchBody.push("Content-Type: application/http\n");
                batchBody.push("Content-Transfer-Encoding: binary\n\n");
                // these are the per-request headers
                const headers = new Headers(init.headers);
                // this is the url of the individual request within the batch
                const reqUrl = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isUrlAbsolute)(url) ? url : (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.combine)(batchQuery.requestBaseUrl, url);
                if (init.method !== "GET") {
                    let method = init.method;
                    if (headers.has("X-HTTP-Method")) {
                        method = headers.get("X-HTTP-Method");
                        headers.delete("X-HTTP-Method");
                    }
                    batchBody.push(`${method} ${reqUrl} HTTP/1.1\n`);
                }
                else {
                    batchBody.push(`${init.method} ${reqUrl} HTTP/1.1\n`);
                }
                // lastly we apply any default headers we need that may not exist
                if (!headers.has("Accept")) {
                    headers.append("Accept", "application/json");
                }
                if (!headers.has("Content-Type")) {
                    headers.append("Content-Type", "application/json;charset=utf-8");
                }
                // write headers into batch body
                headers.forEach((value, name) => {
                    if (headersCopyPattern.test(name)) {
                        batchBody.push(`${name}: ${value}\n`);
                    }
                });
                batchBody.push("\n");
                if (init.body) {
                    batchBody.push(`${init.body}\n\n`);
                }
            }
            if (currentChangeSetId.length > 0) {
                // Close the changeset
                batchBody.push(`--changeset_${currentChangeSetId}--\n\n`);
                currentChangeSetId = "";
            }
            batchBody.push(`--batch_${batchId}--\n`);
            const responses = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(batchQuery, {
                body: batchBody.join(""),
                headers: {
                    "Content-Type": `multipart/mixed; boundary=batch_${batchId}`,
                },
            });
            if (responses.length !== requestsChunk.length) {
                throw Error("Could not properly parse responses to match requests in batch.");
            }
            for (let index = 0; index < responses.length; index++) {
                // resolve the child request's send promise with the parsed response
                requestsChunk[index][3](responses[index]);
            }
        } // end of while (requestsWorkingCopy.length > 0)
        await Promise.all(completePromises).then(() => void (0));
    };
    const register = (instance) => {
        instance.on.init(function () {
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isFunc)(this[RegistrationCompleteSym])) {
                throw Error("This instance is already part of a batch. Please review the docs at https://pnp.github.io/pnpjs/concepts/batching#reuse.");
            }
            // we need to ensure we wait to start execute until all our batch children hit the .send method to be fully registered
            registrationPromises.push(new Promise((resolve) => {
                this[RegistrationCompleteSym] = resolve;
            }));
            return this;
        });
        instance.on.pre(async function (url, init, result) {
            // Do not add to timeline if using BatchNever behavior
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(init.headers, "X-PnP-BatchNever")) {
                // clean up the init operations from the timeline
                // not strictly necessary as none of the logic that uses this should be in the request, but good to keep things tidy
                if (typeof (this[RequestCompleteSym]) === "function") {
                    this[RequestCompleteSym]();
                    delete this[RequestCompleteSym];
                }
                this.using((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.CopyFrom)(refQuery, "replace", (k) => /(init|pre)/i.test(k)));
                return [url, init, result];
            }
            // the entire request will be auth'd - we don't need to run this for each batch request
            this.on.auth.clear();
            // we replace the send function with our batching logic
            this.on.send.replace(async function (url, init) {
                // this is the promise that Queryable will see returned from .emit.send
                const promise = new Promise((resolve) => {
                    // add the request information into the batch
                    requests.push([this, url.toString(), init, resolve]);
                });
                this.log(`[batch:${batchId}] (${(new Date()).getTime()}) Adding request ${init.method} ${url.toString()} to batch.`, 0);
                // we need to ensure we wait to resolve execute until all our batch children have fully completed their request timelines
                completePromises.push(new Promise((resolve) => {
                    this[RequestCompleteSym] = resolve;
                }));
                // indicate that registration of this request is complete
                this[RegistrationCompleteSym]();
                return promise;
            });
            this.on.dispose(function () {
                if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isFunc)(this[RegistrationCompleteSym])) {
                    // if this request is in a batch and caching is in play we need to resolve the registration promises to unblock processing of the batch
                    // because the request will never reach the "send" moment as the result is returned from "pre"
                    this[RegistrationCompleteSym]();
                    // remove the symbol props we added for good hygene
                    delete this[RegistrationCompleteSym];
                }
                if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isFunc)(this[RequestCompleteSym])) {
                    // let things know we are done with this request
                    this[RequestCompleteSym]();
                    delete this[RequestCompleteSym];
                    // there is a code path where you may invoke a batch, say on items.add, whose return
                    // is an object like { data: any, item: IItem }. The expectation from v1 on is `item` in that object
                    // is immediately usable to make additional queries. Without this step when that IItem instance is
                    // created using "this.getById" within IITems.add all of the current observers of "this" are
                    // linked to the IItem instance created (expected), BUT they will be the set of observers setup
                    // to handle the batch, meaning invoking `item` will result in a half batched call that
                    // doesn't really work. To deliver the expected functionality we "reset" the
                    // observers using the original instance, mimicing the behavior had
                    // the IItem been created from that base without a batch involved. We use CopyFrom to ensure
                    // that we maintain the references to the InternalResolve and InternalReject events through
                    // the end of this timeline lifecycle. This works because CopyFrom by design uses Object.keys
                    // which ignores symbol properties.
                    this.using((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.CopyFrom)(refQuery, "replace", (k) => /(auth|pre|send|init|dispose)/i.test(k)));
                }
            });
            return [url, init, result];
        });
        return instance;
    };
    return [register, execute];
}
/**
 * Behavior that blocks batching for the request regardless of "method"
 *
 * This is used for requests to bypass batching methods. Example - Request Digest where we need to get a request-digest inside of a batch.
 * @returns TimelinePipe
 */
function BatchNever() {
    return (instance) => {
        instance.on.pre.prepend(async function (url, init, result) {
            init.headers = { ...init.headers, "X-PnP-BatchNever": "1" };
            return [url, init, result];
        });
        return instance;
    };
}
/**
 * Parses the text body returned by the server from a batch request
 *
 * @param body String body from the server response
 * @returns Parsed response objects
 */
function parseResponse(body) {
    const responses = [];
    const header = "--batchresponse_";
    // Ex. "HTTP/1.1 500 Internal Server Error"
    const statusRegExp = new RegExp("^HTTP/[0-9.]+ +([0-9]+) +(.*)", "i");
    const lines = body.split("\n");
    let state = "batch";
    let status;
    let statusText;
    let headers = {};
    const bodyReader = [];
    for (let i = 0; i < lines.length; ++i) {
        let line = lines[i];
        switch (state) {
            case "batch":
                if (line.substring(0, header.length) === header) {
                    state = "batchHeaders";
                }
                else {
                    if (line.trim() !== "") {
                        throw Error(`Invalid response, line ${i}`);
                    }
                }
                break;
            case "batchHeaders":
                if (line.trim() === "") {
                    state = "status";
                }
                break;
            case "status": {
                const parts = statusRegExp.exec(line);
                if (parts.length !== 3) {
                    throw Error(`Invalid status, line ${i}`);
                }
                status = parseInt(parts[1], 10);
                statusText = parts[2];
                state = "statusHeaders";
                break;
            }
            case "statusHeaders":
                if (line.trim() === "") {
                    state = "body";
                }
                else {
                    const headerParts = line.split(":");
                    if ((headerParts === null || headerParts === void 0 ? void 0 : headerParts.length) === 2) {
                        headers[headerParts[0].trim()] = headerParts[1].trim();
                    }
                }
                break;
            case "body":
                // reset the body reader
                bodyReader.length = 0;
                // this allows us to capture batch bodies that are returned as multi-line (renderListDataAsStream, #2454)
                while (line.substring(0, header.length) !== header) {
                    bodyReader.push(line);
                    line = lines[++i];
                }
                // because we have read the closing --batchresponse_ line, we need to move the line pointer back one
                // so that the logic works as expected either to get the next result or end processing
                i--;
                responses.push(new Response(status === 204 ? null : bodyReader.join(""), { status, statusText, headers }));
                state = "batch";
                headers = {};
                break;
        }
    }
    if (state !== "status") {
        throw Error("Unexpected end of input");
    }
    return responses;
}


/***/ }),

/***/ 7140:
/*!****************************************************!*\
  !*** ./node_modules/@pnp/sp/behaviors/defaults.js ***!
  \****************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   DefaultHeaders: () => (/* binding */ DefaultHeaders),
/* harmony export */   DefaultInit: () => (/* binding */ DefaultInit)
/* harmony export */ });
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _telemetry_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./telemetry.js */ 2630);


function DefaultInit() {
    return (instance) => {
        instance.on.pre(async (url, init, result) => {
            init.cache = "no-cache";
            init.credentials = "same-origin";
            return [url, init, result];
        });
        instance.using((0,_telemetry_js__WEBPACK_IMPORTED_MODULE_1__.Telemetry)(), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.RejectOnError)(), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.ResolveOnData)());
        return instance;
    };
}
function DefaultHeaders() {
    return (instance) => {
        instance
            .using((0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.InjectHeaders)({
            "Accept": "application/json",
            "Content-Type": "application/json;charset=utf-8",
        }));
        return instance;
    };
}


/***/ }),

/***/ 670:
/*!**********************************************************!*\
  !*** ./node_modules/@pnp/sp/behaviors/request-digest.js ***!
  \**********************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   RequestDigest: () => (/* binding */ RequestDigest)
/* harmony export */ });
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../utils/extract-web-url.js */ 2997);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _batching_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../batching.js */ 8018);





function clearExpired(digest) {
    const now = new Date();
    return !(0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.objectDefinedNotNull)(digest) || (now > digest.expiration) ? null : digest;
}
// allows for the caching of digests across all calls which each have their own IDigestInfo wrapper.
const digests = new Map();
function RequestDigest(hook) {
    return (instance) => {
        instance.on.pre(async function (url, init, result) {
            // add the request to the auth moment of the timeline
            this.on.auth(async (url, init) => {
                // eslint-disable-next-line max-len
                if (/get/i.test(init.method) || (init.headers && ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(init.headers, "X-RequestDigest") || (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(init.headers, "Authorization") || (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(init.headers, "X-PnPjs-NoDigest")))) {
                    return [url, init];
                }
                const urlAsString = url.toString();
                const webUrl = (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_2__.extractWebUrl)(urlAsString);
                // do we have one in the cache that is still valid
                // from #2186 we need to always ensure the digest we get isn't expired
                let digest = clearExpired(digests.get(webUrl));
                if (!(0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.objectDefinedNotNull)(digest) && (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isFunc)(hook)) {
                    digest = clearExpired(hook(urlAsString, init));
                }
                if (!(0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.objectDefinedNotNull)(digest)) {
                    digest = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.spPost)((0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.SPQueryable)([this, (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.combine)(webUrl, "_api/contextinfo")]).using((0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.JSONParse)(), (0,_batching_js__WEBPACK_IMPORTED_MODULE_4__.BatchNever)()), {
                        headers: {
                            "Accept": "application/json",
                            "X-PnPjs-NoDigest": "1",
                        },
                    }).then(p => ({
                        expiration: (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.dateAdd)(new Date(), "second", p.FormDigestTimeoutSeconds),
                        value: p.FormDigestValue,
                    }));
                }
                if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.objectDefinedNotNull)(digest)) {
                    // if we got a digest, set it in the headers
                    init.headers = {
                        "X-RequestDigest": digest.value,
                        ...init.headers,
                    };
                    // and cache it for future requests
                    digests.set(webUrl, digest);
                }
                return [url, init];
            });
            return [url, init, result];
        });
        return instance;
    };
}


/***/ }),

/***/ 6793:
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/sp/behaviors/spbrowser.js ***!
  \*****************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* unused harmony export SPBrowser */
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _defaults_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./defaults.js */ 7140);
/* harmony import */ var _request_digest_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./request-digest.js */ 670);




function SPBrowser(props) {
    if ((props === null || props === void 0 ? void 0 : props.baseUrl) && !(0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isUrlAbsolute)(props.baseUrl)) {
        throw Error("SPBrowser props.baseUrl must be absolute when supplied.");
    }
    return (instance) => {
        instance.using((0,_defaults_js__WEBPACK_IMPORTED_MODULE_2__.DefaultHeaders)(), (0,_defaults_js__WEBPACK_IMPORTED_MODULE_2__.DefaultInit)(), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.BrowserFetchWithRetry)(), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.DefaultParse)(), (0,_request_digest_js__WEBPACK_IMPORTED_MODULE_3__.RequestDigest)());
        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isUrlAbsolute)(props === null || props === void 0 ? void 0 : props.baseUrl)) {
            // we want to fix up the url first
            instance.on.pre.prepend(async (url, init, result) => {
                if (!(0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isUrlAbsolute)(url)) {
                    url = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.combine)(props.baseUrl, url);
                }
                return [url, init, result];
            });
        }
        return instance;
    };
}


/***/ }),

/***/ 4243:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/behaviors/spfx.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   SPFx: () => (/* binding */ SPFx)
/* harmony export */ });
/* unused harmony export SPFxToken */
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _defaults_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./defaults.js */ 7140);
/* harmony import */ var _request_digest_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./request-digest.js */ 670);




class SPFxTokenNullOrUndefinedError extends Error {
    constructor(behaviorName) {
        super(`SPFx Context supplied to ${behaviorName} Behavior is null or undefined.`);
    }
    static check(behaviorName, context) {
        if (typeof context === "undefined" || context === null) {
            throw new SPFxTokenNullOrUndefinedError(behaviorName);
        }
    }
}
function SPFxToken(context) {
    SPFxTokenNullOrUndefinedError.check("SPFxToken", context);
    return (instance) => {
        instance.on.auth.replace(async function (url, init) {
            const provider = await context.aadTokenProviderFactory.getTokenProvider();
            const token = await provider.getToken(`${url.protocol}//${url.hostname}`);
            // eslint-disable-next-line @typescript-eslint/dot-notation
            init.headers["Authorization"] = `Bearer ${token}`;
            return [url, init];
        });
        return instance;
    };
}
function SPFx(context) {
    SPFxTokenNullOrUndefinedError.check("SPFx", context);
    return (instance) => {
        instance.using((0,_defaults_js__WEBPACK_IMPORTED_MODULE_2__.DefaultHeaders)(), (0,_defaults_js__WEBPACK_IMPORTED_MODULE_2__.DefaultInit)(), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.BrowserFetchWithRetry)(), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.DefaultParse)(), 
        // remove SPFx Token in default due to issues #2570, #2571
        // SPFxToken(context),
        (0,_request_digest_js__WEBPACK_IMPORTED_MODULE_3__.RequestDigest)((url) => {
            var _a, _b, _c;
            const sameWeb = (new RegExp(`^${(0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.combine)(context.pageContext.web.absoluteUrl, "/_api")}`, "i")).test(url);
            if (sameWeb && ((_b = (_a = context === null || context === void 0 ? void 0 : context.pageContext) === null || _a === void 0 ? void 0 : _a.legacyPageContext) === null || _b === void 0 ? void 0 : _b.formDigestValue)) {
                const creationDateFromDigest = new Date(context.pageContext.legacyPageContext.formDigestValue.split(",")[1]);
                // account for page lifetime in timeout #2304 & others
                // account for tab sleep #2550
                return {
                    value: context.pageContext.legacyPageContext.formDigestValue,
                    expiration: (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.dateAdd)(creationDateFromDigest, "second", ((_c = context.pageContext.legacyPageContext) === null || _c === void 0 ? void 0 : _c.formDigestTimeoutSeconds) - 15 || 1585),
                };
            }
        }));
        // we want to fix up the url first
        instance.on.pre.prepend(async (url, init, result) => {
            if (!(0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isUrlAbsolute)(url)) {
                url = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.combine)(context.pageContext.web.absoluteUrl, url);
            }
            return [url, init, result];
        });
        return instance;
    };
}


/***/ }),

/***/ 2630:
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/sp/behaviors/telemetry.js ***!
  \*****************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Telemetry: () => (/* binding */ Telemetry)
/* harmony export */ });
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);

function Telemetry() {
    return (instance) => {
        instance.on.pre(async function (url, init, result) {
            let clientTag = "PnPCoreJS:4.10.0:";
            // make our best guess based on url to the method called
            const { pathname } = new URL(url);
            // remove anything before the _api as that is potentially PII and we don't care, just want to get the called path to the REST API
            // and we want to modify any (*) calls at the end such as items(3) and items(344) so we just track "items()"
            clientTag = pathname.split("/")
                .filter((v) => !(0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.stringIsNullOrEmpty)(v) && ["_api", "v2.1", "v2.0"].indexOf(v) < 0)
                .map((value, index, arr) => index === arr.length - 1 ? value.replace(/\(.*?$/i, "()") : value[0])
                .join(".");
            if (clientTag.length > 32) {
                clientTag = clientTag.substring(0, 32);
            }
            this.log(`Request Tag: ${clientTag}`, 0);
            init.headers = { ...init.headers, ["X-ClientService-ClientTag"]: clientTag };
            return [url, init, result];
        });
        return instance;
    };
}


/***/ }),

/***/ 5800:
/*!********************************************************!*\
  !*** ./node_modules/@pnp/sp/clientside-pages/funcs.js ***!
  \********************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   getNextOrder: () => (/* binding */ getNextOrder),
/* harmony export */   reindex: () => (/* binding */ reindex)
/* harmony export */ });
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);

/**
 * Gets the next order value 1 based for the provided collection
 *
 * @param collection Collection of orderable things
 */
function getNextOrder(collection) {
    return collection.length < 1 ? 1 : (Math.max.apply(null, collection.map(i => i.order)) + 1);
}
/**
 * Normalizes the order value for all the sections, columns, and controls to be 1 based and stepped (1, 2, 3...)
 *
 * @param collection The collection to normalize
 */
function reindex(collection) {
    for (let i = 0; i < collection.length; i++) {
        collection[i].order = i + 1;
        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(collection[i], "columns")) {
            reindex(collection[i].columns);
        }
        else if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(collection[i], "controls")) {
            reindex(collection[i].controls);
        }
    }
}


/***/ }),

/***/ 5156:
/*!********************************************************!*\
  !*** ./node_modules/@pnp/sp/clientside-pages/index.js ***!
  \********************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _web_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web.js */ 3354);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 7160);




/***/ }),

/***/ 7160:
/*!********************************************************!*\
  !*** ./node_modules/@pnp/sp/clientside-pages/types.js ***!
  \********************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   ClientsidePageFromFile: () => (/* binding */ ClientsidePageFromFile),
/* harmony export */   ClientsideWebpart: () => (/* binding */ ClientsideWebpart),
/* harmony export */   CreateClientsidePage: () => (/* binding */ CreateClientsidePage)
/* harmony export */ });
/* unused harmony exports PromotedState, _ClientsidePage, CanvasSection, CanvasColumn, ColumnControl, ClientsideText */
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _items_types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../items/types.js */ 132);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _lists_types_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../lists/types.js */ 9706);
/* harmony import */ var _utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../utils/odata-url-from.js */ 905);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../utils/extract-web-url.js */ 2997);
/* harmony import */ var _sites_types_js__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../sites/types.js */ 1114);
/* harmony import */ var _funcs_js__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ./funcs.js */ 5800);
/* harmony import */ var _files_web_js__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ../files/web.js */ 6617);
/* harmony import */ var _comments_item_js__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! ../comments/item.js */ 6823);
/* harmony import */ var _batching_js__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(/*! ../batching.js */ 8018);













/**
 * Page promotion state
 */
var PromotedState;
(function (PromotedState) {
    /**
     * Regular client side page
     */
    PromotedState[PromotedState["NotPromoted"] = 0] = "NotPromoted";
    /**
     * Page that will be promoted as news article after publishing
     */
    PromotedState[PromotedState["PromoteOnPublish"] = 1] = "PromoteOnPublish";
    /**
     * Page that is promoted as news article
     */
    PromotedState[PromotedState["Promoted"] = 2] = "Promoted";
})(PromotedState || (PromotedState = {}));
/**
 * Represents the data and methods associated with client side "modern" pages
 */
class _ClientsidePage extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_3__._SPQueryable {
    /**
     * PLEASE DON'T USE THIS CONSTRUCTOR DIRECTLY, thank you 🐇
     */
    constructor(base, path, json, noInit = false, sections = [], commentsDisabled = false) {
        super(base, path);
        this.json = json;
        this.sections = sections;
        this.commentsDisabled = commentsDisabled;
        this._bannerImageDirty = false;
        this._bannerImageThumbnailUrlDirty = false;
        this.parentUrl = "";
        // we need to rebase the url to always be the web url plus the path
        // Queryable handles the correct parsing of the SPInit, and we pull
        // the weburl and combine with the supplied path, which is unique
        // to how ClientsitePages works. This class is a special case.
        this._url = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.combine)((0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_7__.extractWebUrl)(this._url), path);
        // set a default page settings slice
        this._pageSettings = { controlType: 0, pageSettingsSlice: { isDefaultDescription: true, isDefaultThumbnail: true } };
        // set a default layout part
        this._layoutPart = _ClientsidePage.getDefaultLayoutPart();
        if (typeof json !== "undefined" && !noInit) {
            this.fromJSON(json);
        }
    }
    static getDefaultLayoutPart() {
        return {
            dataVersion: "1.4",
            description: "Title Region Description",
            id: "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788",
            instanceId: "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788",
            properties: {
                authorByline: [],
                authors: [],
                layoutType: "FullWidthImage",
                showPublishDate: false,
                showTopicHeader: false,
                textAlignment: "Left",
                title: "",
                topicHeader: "",
                enableGradientEffect: true,
            },
            reservedHeight: 280,
            serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: {}, links: {} },
            title: "Title area",
        };
    }
    get pageLayout() {
        return this.json.PageLayoutType;
    }
    set pageLayout(value) {
        this.json.PageLayoutType = value;
    }
    get bannerImageUrl() {
        return this.json.BannerImageUrl;
    }
    set bannerImageUrl(value) {
        this.setBannerImage(value);
    }
    get thumbnailUrl() {
        return this._pageSettings.pageSettingsSlice.isDefaultThumbnail ? this.json.BannerImageUrl : this.json.BannerThumbnailUrl;
    }
    set thumbnailUrl(value) {
        this.json.BannerThumbnailUrl = value;
        this._bannerImageThumbnailUrlDirty = true;
        this._pageSettings.pageSettingsSlice.isDefaultThumbnail = false;
    }
    get topicHeader() {
        return (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.objectDefinedNotNull)(this.json.TopicHeader) ? this.json.TopicHeader : "";
    }
    set topicHeader(value) {
        this.json.TopicHeader = value;
        this._layoutPart.properties.topicHeader = value;
        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.stringIsNullOrEmpty)(value)) {
            this.showTopicHeader = false;
        }
    }
    get title() {
        return this.json.Title;
    }
    set title(value) {
        this.json.Title = value;
        this._layoutPart.properties.title = value;
    }
    get reservedHeight() {
        return this._layoutPart.reservedHeight;
    }
    set reservedHeight(value) {
        this._layoutPart.reservedHeight = value;
    }
    get description() {
        return this.json.Description;
    }
    set description(value) {
        if (!(0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.stringIsNullOrEmpty)(value) && value.length > 255) {
            throw Error("Modern Page description is limited to 255 chars.");
        }
        this.json.Description = value;
        if (!(0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(this._pageSettings, "htmlAttributes")) {
            this._pageSettings.htmlAttributes = [];
        }
        if (this._pageSettings.htmlAttributes.indexOf("modifiedDescription") < 0) {
            this._pageSettings.htmlAttributes.push("modifiedDescription");
        }
        this._pageSettings.pageSettingsSlice.isDefaultDescription = false;
    }
    get layoutType() {
        return this._layoutPart.properties.layoutType;
    }
    set layoutType(value) {
        this._layoutPart.properties.layoutType = value;
    }
    get headerTextAlignment() {
        return this._layoutPart.properties.textAlignment;
    }
    set headerTextAlignment(value) {
        this._layoutPart.properties.textAlignment = value;
    }
    get showTopicHeader() {
        return this._layoutPart.properties.showTopicHeader;
    }
    set showTopicHeader(value) {
        this._layoutPart.properties.showTopicHeader = value;
    }
    get showPublishDate() {
        return this._layoutPart.properties.showPublishDate;
    }
    set showPublishDate(value) {
        this._layoutPart.properties.showPublishDate = value;
    }
    get hasVerticalSection() {
        return this.sections.findIndex(s => s.layoutIndex === 2) > -1;
    }
    get authorByLine() {
        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.isArray)(this._layoutPart.properties.authorByline) && this._layoutPart.properties.authorByline.length > 0) {
            return this._layoutPart.properties.authorByline[0];
        }
        return null;
    }
    get verticalSection() {
        if (this.hasVerticalSection) {
            return this.addVerticalSection();
        }
        return null;
    }
    /**
     * Add a section to this page
     */
    addSection() {
        const section = new CanvasSection(this, (0,_funcs_js__WEBPACK_IMPORTED_MODULE_9__.getNextOrder)(this.sections), 1);
        this.sections.push(section);
        return section;
    }
    /**
     * Add a section to this page
     */
    addVerticalSection() {
        // we can only have one vertical section so we find it if it exists
        const sectionIndex = this.sections.findIndex(s => s.layoutIndex === 2);
        if (sectionIndex > -1) {
            return this.sections[sectionIndex];
        }
        const section = new CanvasSection(this, (0,_funcs_js__WEBPACK_IMPORTED_MODULE_9__.getNextOrder)(this.sections), 2);
        this.sections.push(section);
        return section;
    }
    /**
     * Loads this instance from the appropriate JSON data
     *
     * @param pageData JSON data to load (replaces any existing data)
     */
    fromJSON(pageData) {
        this.json = pageData;
        const canvasControls = JSON.parse(pageData.CanvasContent1);
        const layouts = JSON.parse(pageData.LayoutWebpartsContent);
        if (layouts && layouts.length > 0) {
            this._layoutPart = layouts[0];
        }
        this.setControls(canvasControls);
        return this;
    }
    /**
     * Loads this page's content from the server
     */
    async load() {
        const item = await this.getItem("Id", "CommentsDisabled");
        const pageData = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.SPQueryable)(this, `_api/sitepages/pages(${item.Id})`)();
        this.commentsDisabled = item.CommentsDisabled;
        return this.fromJSON(pageData);
    }
    /**
     * Persists the content changes (sections, columns, and controls) [does not work with batching]
     *
     * @param publish If true the page is published, if false the changes are persisted to SharePoint but not published [Default: true]
     */
    async save(publish = true) {
        if (this.json.Id === null) {
            throw Error("The id for this page is null. If you want to create a new page, please use ClientSidePage.Create");
        }
        const previewPartialUrl = "_layouts/15/getpreview.ashx";
        // If new banner image, and banner image url is not in getpreview.ashx format
        if (this._bannerImageDirty && !this.bannerImageUrl.includes(previewPartialUrl)) {
            const serverRelativePath = this.bannerImageUrl;
            let imgInfo;
            let webUrl;
            const web = (0,_webs_types_js__WEBPACK_IMPORTED_MODULE_6__.Web)(this);
            const [batch, execute] = (0,_batching_js__WEBPACK_IMPORTED_MODULE_12__.createBatch)(web);
            web.using(batch);
            web.getFileByServerRelativePath(serverRelativePath.replace(/%20/ig, " "))
                .select("ListId", "WebId", "UniqueId", "Name", "SiteId")().then(r1 => imgInfo = r1);
            web.select("Url")().then(r2 => webUrl = r2.Url);
            // we know the .then calls above will run before execute resolves, ensuring the vars are set
            await execute();
            const f = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.SPQueryable)(webUrl, previewPartialUrl);
            f.query.set("guidSite", `${imgInfo.SiteId}`);
            f.query.set("guidWeb", `${imgInfo.WebId}`);
            f.query.set("guidFile", `${imgInfo.UniqueId}`);
            this.bannerImageUrl = f.toRequestUrl();
            if (!(0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.objectDefinedNotNull)(this._layoutPart.serverProcessedContent)) {
                this._layoutPart.serverProcessedContent = {};
            }
            this._layoutPart.serverProcessedContent.imageSources = { imageSource: serverRelativePath };
            if (!(0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.objectDefinedNotNull)(this._layoutPart.serverProcessedContent.customMetadata)) {
                this._layoutPart.serverProcessedContent.customMetadata = {};
            }
            this._layoutPart.serverProcessedContent.customMetadata.imageSource = {
                listId: imgInfo.ListId,
                siteId: imgInfo.SiteId,
                uniqueId: imgInfo.UniqueId,
                webId: imgInfo.WebId,
            };
            this._layoutPart.properties.webId = imgInfo.WebId;
            this._layoutPart.properties.siteId = imgInfo.SiteId;
            this._layoutPart.properties.listId = imgInfo.ListId;
            this._layoutPart.properties.uniqueId = imgInfo.UniqueId;
        }
        // we try and check out the page for the user
        if (!this.json.IsPageCheckedOutToCurrentUser) {
            await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.spPost)(ClientsidePage(this, `_api/sitepages/pages(${this.json.Id})/checkoutpage`));
        }
        // create the body for the save request
        let saveBody = {
            AuthorByline: this.json.AuthorByline || [],
            CanvasContent1: this.getCanvasContent1(),
            Description: this.description,
            LayoutWebpartsContent: this.getLayoutWebpartsContent(),
            Title: this.title,
            TopicHeader: this.topicHeader,
            BannerImageUrl: this.bannerImageUrl,
        };
        if (this._bannerImageDirty || this._bannerImageThumbnailUrlDirty) {
            const bannerImageUrlValue = this._bannerImageThumbnailUrlDirty ? this.thumbnailUrl : this.bannerImageUrl;
            saveBody = {
                BannerImageUrl: bannerImageUrlValue,
                ...saveBody,
            };
        }
        const updater = ClientsidePage(this, `_api/sitepages/pages(${this.json.Id})/savepage`);
        await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.spPost)(updater, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.headers)({ "if-match": "*" }, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)(saveBody)));
        let r = true;
        if (publish) {
            r = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.spPost)(ClientsidePage(this, `_api/sitepages/pages(${this.json.Id})/publish`));
            if (r) {
                this.json.IsPageCheckedOutToCurrentUser = false;
            }
        }
        this._bannerImageDirty = false;
        this._bannerImageThumbnailUrlDirty = false;
        // we need to ensure we reload from the latest data to ensure all urls are updated and current in the object (expecially for new pages)
        await this.load();
        return r;
    }
    /**
     * Discards the checkout of this page
     */
    async discardPageCheckout() {
        if (this.json.Id === null) {
            throw Error("The id for this page is null. If you want to create a new page, please use ClientSidePage.Create");
        }
        const d = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.spPost)(ClientsidePage(this, `_api/sitepages/pages(${this.json.Id})/discardPage`));
        this.fromJSON(d);
    }
    /**
     * Promotes this page as a news item
     */
    async promoteToNews() {
        return this.promoteNewsImpl("promoteToNews");
    }
    // API is currently broken on server side
    // public async demoteFromNews(): Promise<boolean> {
    //     return this.promoteNewsImpl("demoteFromNews");
    // }
    /**
     * Finds a control by the specified instance id
     *
     * @param id Instance id of the control to find
     */
    findControlById(id) {
        return this.findControl((c) => c.id === id);
    }
    /**
     * Finds a control within this page's control tree using the supplied predicate
     *
     * @param predicate Takes a control and returns true or false, if true that control is returned by findControl
     */
    findControl(predicate) {
        // check all sections
        for (let i = 0; i < this.sections.length; i++) {
            // check all columns
            for (let j = 0; j < this.sections[i].columns.length; j++) {
                // check all controls
                for (let k = 0; k < this.sections[i].columns[j].controls.length; k++) {
                    // check to see if the predicate likes this control
                    if (predicate(this.sections[i].columns[j].controls[k])) {
                        return this.sections[i].columns[j].controls[k];
                    }
                }
            }
        }
        // we found nothing so give nothing back
        return null;
    }
    /**
     * Creates a new page with all of the content copied from this page
     *
     * @param web The web where we will create the copy
     * @param pageName The file name of the new page
     * @param title The title of the new page
     * @param publish If true the page will be published (Default: true)
     */
    async copy(web, pageName, title, publish = true, promotedState) {
        const page = await CreateClientsidePage(web, pageName, title, this.pageLayout, promotedState);
        return this.copyTo(page, publish);
    }
    /**
     * Copies the content from this page to the supplied page instance NOTE: fully overwriting any previous content!!
     *
     * @param page Page whose content we replace with this page's content
     * @param publish If true the page will be published after the copy, if false it will be saved but left unpublished (Default: true)
     */
    async copyTo(page, publish = true) {
        // we know the method is on the class - but it is protected so not part of the interface
        page.setControls(this.getControls());
        // copy page properties
        if (this._layoutPart.properties) {
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(this._layoutPart.properties, "topicHeader")) {
                page.topicHeader = this._layoutPart.properties.topicHeader;
            }
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(this._layoutPart.properties, "imageSourceType")) {
                page._layoutPart.properties.imageSourceType = this._layoutPart.properties.imageSourceType;
            }
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(this._layoutPart.properties, "layoutType")) {
                page._layoutPart.properties.layoutType = this._layoutPart.properties.layoutType;
            }
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(this._layoutPart.properties, "textAlignment")) {
                page._layoutPart.properties.textAlignment = this._layoutPart.properties.textAlignment;
            }
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(this._layoutPart.properties, "showTopicHeader")) {
                page._layoutPart.properties.showTopicHeader = this._layoutPart.properties.showTopicHeader;
            }
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(this._layoutPart.properties, "showPublishDate")) {
                page._layoutPart.properties.showPublishDate = this._layoutPart.properties.showPublishDate;
            }
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(this._layoutPart.properties, "enableGradientEffect")) {
                page._layoutPart.properties.enableGradientEffect = this._layoutPart.properties.enableGradientEffect;
            }
        }
        // we need to do some work to set the banner image url in the copied page
        if (!(0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.stringIsNullOrEmpty)(this.json.BannerImageUrl)) {
            // use a URL to parse things for us
            const url = new URL(this.json.BannerImageUrl);
            // helper function to translate the guid strings into properly formatted guids with dashes
            const makeGuid = (s) => s.replace(/(.{8})(.{4})(.{4})(.{4})(.{12})/g, "$1-$2-$3-$4-$5");
            // protect against errors because the serverside impl has changed, we'll just skip
            if (url.searchParams.has("guidSite") && url.searchParams.has("guidWeb") && url.searchParams.has("guidFile")) {
                const guidSite = makeGuid(url.searchParams.get("guidSite"));
                const guidWeb = makeGuid(url.searchParams.get("guidWeb"));
                const guidFile = makeGuid(url.searchParams.get("guidFile"));
                const site = (0,_sites_types_js__WEBPACK_IMPORTED_MODULE_8__.Site)(this);
                const id = await site.select("Id")();
                // the site guid must match the current site's guid or we are unable to set the image
                if (id.Id === guidSite) {
                    const openWeb = await site.openWebById(guidWeb);
                    const file = await openWeb.web.getFileById(guidFile).select("ServerRelativeUrl")();
                    const props = {};
                    if (this._layoutPart.properties) {
                        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(this._layoutPart.properties, "translateX")) {
                            props.translateX = this._layoutPart.properties.translateX;
                        }
                        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(this._layoutPart.properties, "translateY")) {
                            props.translateY = this._layoutPart.properties.translateY;
                        }
                        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(this._layoutPart.properties, "imageSourceType")) {
                            props.imageSourceType = this._layoutPart.properties.imageSourceType;
                        }
                        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(this._layoutPart.properties, "altText")) {
                            props.altText = this._layoutPart.properties.altText;
                        }
                    }
                    page.setBannerImage(file.ServerRelativeUrl, props);
                }
            }
        }
        await page.save(publish);
        return page;
    }
    /**
     * Sets the modern page banner image
     *
     * @param url Url of the image to display
     * @param altText Alt text to describe the image
     * @param bannerProps Additional properties to control display of the banner
     */
    setBannerImage(url, props) {
        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.isUrlAbsolute)(url)) {
            // do our best to make this a server relative url by removing the x.sharepoint.com part
            url = url.replace(/^https?:\/\/[a-z0-9.]*?\.[a-z]{2,3}\//i, "/");
        }
        this.json.BannerImageUrl = url;
        // update serverProcessedContent (page behavior change 2021-Oct-13)
        this._layoutPart.serverProcessedContent = { imageSources: { imageSource: url } };
        this._bannerImageDirty = true;
        /*
            setting the banner image resets the thumbnail image (matching UI functionality)
            but if the thumbnail is dirty they are likely trying to set them both to
            different values, so we allow that here.
            Also allows the banner image to be updated safely with the calculated one in save()
        */
        if (!this._bannerImageThumbnailUrlDirty) {
            this.thumbnailUrl = url;
            this._pageSettings.pageSettingsSlice.isDefaultThumbnail = true;
        }
        // this seems to always be true, so default
        this._layoutPart.properties.imageSourceType = 2;
        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.objectDefinedNotNull)(props)) {
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(props, "translateX")) {
                this._layoutPart.properties.translateX = props.translateX;
            }
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(props, "translateY")) {
                this._layoutPart.properties.translateY = props.translateY;
            }
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(props, "imageSourceType")) {
                this._layoutPart.properties.imageSourceType = props.imageSourceType;
            }
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(props, "altText")) {
                this._layoutPart.properties.altText = props.altText;
            }
        }
    }
    /**
     * Sets the banner image url from an external source. You must call save to persist the changes
     *
     * @param url absolute url of the external file
     * @param props optional set of properties to control display of the banner image
     */
    async setBannerImageFromExternalUrl(url, props) {
        // validate and parse our input url
        const fileUrl = new URL(url);
        // get our page name without extension, used as a folder name when creating the file
        const pageName = this.json.FileName.replace(/\.[^/.]+$/, "");
        // get the filename we will use
        const filename = fileUrl.pathname.split(/[\\/]/i).pop();
        const request = ClientsidePage(this, "_api/sitepages/AddImageFromExternalUrl");
        request.query.set("imageFileName", `'${filename}'`);
        request.query.set("pageName", `'${pageName}'`);
        request.query.set("externalUrl", `'${url}'`);
        request.select("ServerRelativeUrl");
        const result = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.spPost)(request);
        // set with the newly created relative url
        this.setBannerImage(result.ServerRelativeUrl, props);
    }
    /**
     * Sets the authors for this page from the supplied list of user integer ids
     *
     * @param authorId The integer id of the user to set as the author
     */
    async setAuthorById(authorId) {
        const userLoginData = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.SPCollection)([this, (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_7__.extractWebUrl)(this.toUrl())], "/_api/web/siteusers")
            .filter(`Id eq ${authorId}`)
            .select("LoginName")();
        if (userLoginData.length < 1) {
            throw Error(`Could not find user with id ${authorId}.`);
        }
        return this.setAuthorByLoginName(userLoginData[0].LoginName);
    }
    /**
     * Sets the authors for this page from the supplied list of user integer ids
     *
     * @param authorLoginName The login name of the user (ex: i:0#.f|membership|name@tenant.com)
     */
    async setAuthorByLoginName(authorLoginName) {
        const userLoginData = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.SPCollection)([this, (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_7__.extractWebUrl)(this.toUrl())], "/_api/web/siteusers")
            .filter(`LoginName eq '${authorLoginName}'`)
            .select("UserPrincipalName", "Title")();
        if (userLoginData.length < 1) {
            throw Error(`Could not find user with login name '${authorLoginName}'.`);
        }
        this.json.AuthorByline = [userLoginData[0].UserPrincipalName];
        this._layoutPart.properties.authorByline = [userLoginData[0].UserPrincipalName];
        this._layoutPart.properties.authors = [{
                id: authorLoginName,
                name: userLoginData[0].Title,
                role: "",
                upn: userLoginData[0].UserPrincipalName,
            }];
    }
    /**
     * Gets the list item associated with this clientside page
     *
     * @param selects Specific set of fields to include when getting the item
     */
    async getItem(...selects) {
        const initer = ClientsidePage(this, "/_api/lists/EnsureClientRenderedSitePagesLibrary").select("EnableModeration", "EnableMinorVersions", "Id");
        const listData = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.spPost)(initer);
        const item = (0,_lists_types_js__WEBPACK_IMPORTED_MODULE_4__.List)([this, listData["odata.id"]]).items.getById(this.json.Id);
        const itemData = await item.select(...selects)();
        return Object.assign((0,_items_types_js__WEBPACK_IMPORTED_MODULE_2__.Item)([this, (0,_utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_5__.odataUrlFrom)(itemData)]), itemData);
    }
    /**
         * Recycle this page
         */
    async recycle() {
        const item = await this.getItem();
        await item.recycle();
    }
    /**
     * Delete this page
     */
    async delete() {
        const item = await this.getItem();
        await item.delete();
    }
    /**
     * Schedules a page for publishing
     *
     * @param publishDate Date to publish the item
     * @returns Version which was scheduled to be published
     */
    async schedulePublish(publishDate) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.spPost)(ClientsidePage(this, `_api/sitepages/pages(${this.json.Id})/SchedulePublish`), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({
            sitePage: { PublishStartDate: publishDate },
        }));
    }
    /**
     * Saves a copy of this page as a template in this library's Templates folder
     *
     * @param publish If true the template is published, false the template is not published (default: true)
     * @returns IClientsidePage instance representing the new template page
     */
    async saveAsTemplate(publish = true) {
        const data = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.spPost)(ClientsidePage(this, `_api/sitepages/pages(${this.json.Id})/SavePageAsTemplate`));
        const page = ClientsidePage(this, null, data);
        page.title = this.title;
        await page.save(publish);
        return page;
    }
    /**
     * Share this Page's Preview content by Email
     *
     * @param emails Set of emails to which the preview is shared
     * @param message The message to include
     * @returns void
     */
    share(emails, message) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.spPost)(ClientsidePage(this, "_api/SP.Publishing.RichSharing/SharePageByEmail"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({
            recipientEmails: emails,
            message,
            url: this.json.AbsoluteUrl,
        }));
    }
    getCanvasContent1() {
        return JSON.stringify(this.getControls());
    }
    getLayoutWebpartsContent() {
        if (this._layoutPart) {
            return JSON.stringify([this._layoutPart]);
        }
        else {
            return JSON.stringify(null);
        }
    }
    setControls(controls) {
        // reset the sections
        this.sections = [];
        if (controls && controls.length) {
            for (let i = 0; i < controls.length; i++) {
                // if no control type is present this is a column which we give type 0 to let us process it
                const controlType = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(controls[i], "controlType") ? controls[i].controlType : 0;
                switch (controlType) {
                    case 0:
                        // empty canvas column or page settings
                        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(controls[i], "pageSettingsSlice")) {
                            this._pageSettings = controls[i];
                        }
                        else {
                            // we have an empty column
                            this.mergeColumnToTree(new CanvasColumn(controls[i]));
                        }
                        break;
                    case 3: {
                        const part = new ClientsideWebpart(controls[i]);
                        this.mergePartToTree(part, part.data.position);
                        break;
                    }
                    case 4: {
                        const textData = controls[i];
                        const text = new ClientsideText(textData.innerHTML, textData);
                        this.mergePartToTree(text, text.data.position);
                        break;
                    }
                }
            }
            (0,_funcs_js__WEBPACK_IMPORTED_MODULE_9__.reindex)(this.sections);
        }
    }
    getControls() {
        // reindex things
        (0,_funcs_js__WEBPACK_IMPORTED_MODULE_9__.reindex)(this.sections);
        // rollup the control changes
        const canvasData = [];
        this.sections.forEach(section => {
            section.columns.forEach(column => {
                if (column.controls.length < 1) {
                    // empty column
                    canvasData.push({
                        displayMode: column.data.displayMode,
                        emphasis: this.getEmphasisObj(section.emphasis),
                        position: column.data.position,
                    });
                }
                else {
                    column.controls.forEach(control => {
                        control.data.emphasis = this.getEmphasisObj(section.emphasis);
                        canvasData.push(this.specialSaveHandling(control).data);
                    });
                }
            });
        });
        canvasData.push(this._pageSettings);
        return canvasData;
    }
    getEmphasisObj(value) {
        if (value < 1 || value > 3) {
            return {};
        }
        return { zoneEmphasis: value };
    }
    async promoteNewsImpl(method) {
        if (this.json.Id === null) {
            throw Error("The id for this page is null.");
        }
        // per bug #858 if we promote before we have ever published the last published date will
        // forever not be updated correctly in the modern news web part. Because this will affect very
        // few folks we just go ahead and publish for them here as that is likely what they intended.
        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.stringIsNullOrEmpty)(this.json.VersionInfo.LastVersionCreatedBy)) {
            const lastPubData = new Date(this.json.VersionInfo.LastVersionCreated);
            // no modern page should reasonable be published before the year 2000 :)
            if (lastPubData.getFullYear() < 2000) {
                await this.save(true);
            }
        }
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.spPost)(ClientsidePage(this, `_api/sitepages/pages(${this.json.Id})/${method}`));
    }
    /**
     * Merges the control into the tree of sections and columns for this page
     *
     * @param control The control to merge
     */
    mergePartToTree(control, positionData) {
        var _a, _b, _c;
        let column = null;
        let sectionFactor = 12;
        let sectionIndex = 0;
        let zoneIndex = 0;
        let layoutIndex = 1;
        // handle case where we don't have position data (shouldn't happen?)
        if (positionData) {
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(positionData, "zoneIndex")) {
                zoneIndex = positionData.zoneIndex;
            }
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(positionData, "sectionIndex")) {
                sectionIndex = positionData.sectionIndex;
            }
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(positionData, "sectionFactor")) {
                sectionFactor = positionData.sectionFactor;
            }
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(positionData, "layoutIndex")) {
                layoutIndex = positionData.layoutIndex;
            }
        }
        const zoneEmphasis = (_c = (_b = (_a = control.data) === null || _a === void 0 ? void 0 : _a.emphasis) === null || _b === void 0 ? void 0 : _b.zoneEmphasis) !== null && _c !== void 0 ? _c : 0;
        const section = this.getOrCreateSection(zoneIndex, layoutIndex, zoneEmphasis);
        const columns = section.columns.filter(c => c.order === sectionIndex);
        if (columns.length < 1) {
            column = section.addColumn(sectionFactor, layoutIndex);
        }
        else {
            column = columns[0];
        }
        control.column = column;
        column.addControl(control);
    }
    /**
     * Merges the supplied column into the tree
     *
     * @param column Column to merge
     * @param position The position data for the column
     */
    mergeColumnToTree(column) {
        var _a, _b;
        const order = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(column.data, "position") && (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(column.data.position, "zoneIndex") ? column.data.position.zoneIndex : 0;
        const layoutIndex = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(column.data, "position") && (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(column.data.position, "layoutIndex") ? column.data.position.layoutIndex : 1;
        const section = this.getOrCreateSection(order, layoutIndex, ((_b = (_a = column.data) === null || _a === void 0 ? void 0 : _a.emphasis) === null || _b === void 0 ? void 0 : _b.zoneEmphasis) || 0);
        column.section = section;
        section.columns.push(column);
    }
    /**
     * Handle the logic to get or create a section based on the supplied order and layoutIndex
     *
     * @param order Section order
     * @param layoutIndex Layout Index (1 === normal, 2 === vertical section)
     * @param emphasis The section emphasis
     */
    getOrCreateSection(order, layoutIndex, emphasis) {
        let section = null;
        const sections = this.sections.filter(s => s.order === order && s.layoutIndex === layoutIndex);
        if (sections.length < 1) {
            section = layoutIndex === 2 ? this.addVerticalSection() : this.addSection();
            section.order = order;
            section.emphasis = emphasis;
        }
        else {
            section = sections[0];
        }
        return section;
    }
    /**
     * Based on issue #1690 we need to take special case actions to ensure some things
     * can be saved properly without breaking existing pages.
     *
     * @param control The control we are ensuring is "ready" to be saved
     */
    specialSaveHandling(control) {
        var _a, _b, _c;
        // this is to handle the special case in issue #1690
        // must ensure that searchablePlainTexts values have < replaced with &lt; in links web part
        // For #2561 need to process for code snippet webpart and any control && (<any>control).data.webPartId === "c70391ea-0b10-4ee9-b2b4-006d3fcad0cd"
        if (control.data.controlType === 3) {
            const texts = ((_c = (_b = (_a = control.data) === null || _a === void 0 ? void 0 : _a.webPartData) === null || _b === void 0 ? void 0 : _b.serverProcessedContent) === null || _c === void 0 ? void 0 : _c.searchablePlainTexts) || null;
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.objectDefinedNotNull)(texts)) {
                const keys = Object.getOwnPropertyNames(texts);
                for (let i = 0; i < keys.length; i++) {
                    texts[keys[i]] = texts[keys[i]].replace(/</ig, "&lt;");
                    control.data.webPartData.serverProcessedContent.searchablePlainTexts = texts;
                }
            }
        }
        return control;
    }
}
/**
 * Invokable factory for IClientSidePage instances
 */
const ClientsidePage = (base, path, json, noInit = false, sections = [], commentsDisabled = false) => {
    return new _ClientsidePage(base, path, json, noInit, sections, commentsDisabled);
};
/**
 * Loads a client side page from the supplied IFile instance
 *
 * @param file Source IFile instance
 */
const ClientsidePageFromFile = async (file) => {
    const item = await file.getItem();
    const page = ClientsidePage([file, (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_7__.extractWebUrl)(file.toUrl())], "", { Id: item.Id }, true);
    return page.load();
};
/**
 * Creates a new client side page
 *
 * @param web The web or list
 * @param pageName The name of the page (filename)
 * @param title The page's title
 * @param PageLayoutType Layout to use when creating the page
 */
const CreateClientsidePage = async (web, pageName, title, PageLayoutType = "Article", promotedState = 0) => {
    // patched because previously we used the full page name with the .aspx at the end
    // this allows folk's existing code to work after the re-write to the new API
    pageName = pageName.replace(/\.aspx$/i, "");
    // initialize the page, at this point a checked-out page with a junk filename will be created.
    const pageInitData = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.spPost)(ClientsidePage(web, "_api/sitepages/pages"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({
        PageLayoutType,
        PromotedState: promotedState,
    }));
    // now we can init our page with the save data
    const newPage = ClientsidePage(web, "", pageInitData);
    newPage.title = pageName;
    await newPage.save(false);
    newPage.title = title;
    return newPage;
};
class CanvasSection {
    constructor(page, order, layoutIndex, columns = [], _emphasis = 0) {
        this.page = page;
        this.columns = columns;
        this._emphasis = _emphasis;
        this._memId = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.getGUID)();
        this._order = order;
        this._layoutIndex = layoutIndex;
    }
    get order() {
        return this._order;
    }
    set order(value) {
        this._order = value;
        for (let i = 0; i < this.columns.length; i++) {
            this.columns[i].data.position.zoneIndex = value;
        }
    }
    get layoutIndex() {
        return this._layoutIndex;
    }
    set layoutIndex(value) {
        this._layoutIndex = value;
        for (let i = 0; i < this.columns.length; i++) {
            this.columns[i].data.position.layoutIndex = value;
        }
    }
    /**
     * Default column (this.columns[0]) for this section
     */
    get defaultColumn() {
        if (this.columns.length < 1) {
            this.addColumn(12);
        }
        return this.columns[0];
    }
    /**
     * Adds a new column to this section
     */
    addColumn(factor, layoutIndex = this.layoutIndex) {
        const column = new CanvasColumn();
        column.section = this;
        column.data.position.zoneIndex = this.order;
        column.data.position.layoutIndex = layoutIndex;
        column.data.position.sectionFactor = factor;
        column.order = (0,_funcs_js__WEBPACK_IMPORTED_MODULE_9__.getNextOrder)(this.columns);
        this.columns.push(column);
        return column;
    }
    /**
     * Adds a control to the default column for this section
     *
     * @param control Control to add to the default column
     */
    addControl(control) {
        this.defaultColumn.addControl(control);
        return this;
    }
    get emphasis() {
        return this._emphasis;
    }
    set emphasis(value) {
        this._emphasis = value;
    }
    /**
     * Removes this section and all contained columns and controls from the collection
     */
    remove() {
        this.page.sections = this.page.sections.filter(section => section._memId !== this._memId);
        (0,_funcs_js__WEBPACK_IMPORTED_MODULE_9__.reindex)(this.page.sections);
    }
}
class CanvasColumn {
    constructor(json = JSON.parse(JSON.stringify(CanvasColumn.Default)), controls = []) {
        this.json = json;
        this.controls = controls;
        this._section = null;
        this._memId = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.getGUID)();
    }
    get data() {
        return this.json;
    }
    get section() {
        return this._section;
    }
    set section(section) {
        this._section = section;
    }
    get order() {
        return this.data.position.sectionIndex;
    }
    set order(value) {
        this.data.position.sectionIndex = value;
        for (let i = 0; i < this.controls.length; i++) {
            this.controls[i].data.position.zoneIndex = this.data.position.zoneIndex;
            this.controls[i].data.position.layoutIndex = this.data.position.layoutIndex;
            this.controls[i].data.position.sectionIndex = value;
        }
    }
    get factor() {
        return this.data.position.sectionFactor;
    }
    set factor(value) {
        this.data.position.sectionFactor = value;
    }
    addControl(control) {
        control.column = this;
        this.controls.push(control);
        return this;
    }
    getControl(index) {
        return this.controls[index];
    }
    remove() {
        this.section.columns = this.section.columns.filter(column => column._memId !== this._memId);
        (0,_funcs_js__WEBPACK_IMPORTED_MODULE_9__.reindex)(this.section.columns);
    }
}
CanvasColumn.Default = {
    controlType: 0,
    displayMode: 2,
    emphasis: {},
    position: {
        layoutIndex: 1,
        sectionFactor: 12,
        sectionIndex: 1,
        zoneIndex: 1,
    },
};
class ColumnControl {
    constructor(json) {
        this.json = json;
    }
    get id() {
        return this.json.id;
    }
    get data() {
        return this.json;
    }
    get column() {
        return this._column;
    }
    set column(value) {
        this._column = value;
        this.onColumnChange(this._column);
    }
    remove() {
        this.column.controls = this.column.controls.filter(control => control.id !== this.id);
        (0,_funcs_js__WEBPACK_IMPORTED_MODULE_9__.reindex)(this.column.controls);
    }
    setData(data) {
        this.json = data;
    }
}
class ClientsideText extends ColumnControl {
    constructor(text, json = JSON.parse(JSON.stringify(ClientsideText.Default))) {
        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.stringIsNullOrEmpty)(json.id)) {
            json.id = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.getGUID)();
            json.anchorComponentId = json.id;
        }
        super(json);
        this.text = text;
    }
    get text() {
        return this.data.innerHTML;
    }
    set text(value) {
        this.data.innerHTML = value;
    }
    get order() {
        return this.data.position.controlIndex;
    }
    set order(value) {
        this.data.position.controlIndex = value;
    }
    onColumnChange(col) {
        this.data.position.sectionFactor = col.factor;
        this.data.position.controlIndex = (0,_funcs_js__WEBPACK_IMPORTED_MODULE_9__.getNextOrder)(col.controls);
        this.data.position.zoneIndex = col.data.position.zoneIndex;
        this.data.position.sectionIndex = col.order;
        this.data.position.layoutIndex = col.data.position.layoutIndex;
    }
}
ClientsideText.Default = {
    addedFromPersistedData: false,
    anchorComponentId: "",
    controlType: 4,
    displayMode: 2,
    editorType: "CKEditor",
    emphasis: {},
    id: "",
    innerHTML: "",
    position: {
        controlIndex: 1,
        layoutIndex: 1,
        sectionFactor: 12,
        sectionIndex: 1,
        zoneIndex: 1,
    },
};
class ClientsideWebpart extends ColumnControl {
    constructor(json = JSON.parse(JSON.stringify(ClientsideWebpart.Default))) {
        super(json);
    }
    static fromComponentDef(definition) {
        const part = new ClientsideWebpart();
        part.import(definition);
        return part;
    }
    get title() {
        return this.data.webPartData.title;
    }
    set title(value) {
        this.data.webPartData.title = value;
    }
    get description() {
        return this.data.webPartData.description;
    }
    set description(value) {
        this.data.webPartData.description = value;
    }
    get order() {
        return this.data.position.controlIndex;
    }
    set order(value) {
        this.data.position.controlIndex = value;
    }
    get height() {
        return this.data.reservedHeight;
    }
    set height(value) {
        this.data.reservedHeight = value;
    }
    get width() {
        return this.data.reservedWidth;
    }
    set width(value) {
        this.data.reservedWidth = value;
    }
    get dataVersion() {
        return this.data.webPartData.dataVersion;
    }
    set dataVersion(value) {
        this.data.webPartData.dataVersion = value;
    }
    setProperties(properties) {
        this.data.webPartData.properties = {
            ...this.data.webPartData.properties,
            ...properties,
        };
        return this;
    }
    getProperties() {
        return this.data.webPartData.properties;
    }
    setServerProcessedContent(properties) {
        this.data.webPartData.serverProcessedContent = {
            ...this.data.webPartData.serverProcessedContent,
            ...properties,
        };
        return this;
    }
    getServerProcessedContent() {
        return this.data.webPartData.serverProcessedContent;
    }
    onColumnChange(col) {
        this.data.position.sectionFactor = col.factor;
        this.data.position.controlIndex = (0,_funcs_js__WEBPACK_IMPORTED_MODULE_9__.getNextOrder)(col.controls);
        this.data.position.zoneIndex = col.data.position.zoneIndex;
        this.data.position.sectionIndex = col.data.position.sectionIndex;
        this.data.position.layoutIndex = col.data.position.layoutIndex;
    }
    import(component) {
        const id = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.getGUID)();
        const componendId = component.Id.replace(/^\{|\}$/g, "").toLowerCase();
        const manifest = JSON.parse(component.Manifest);
        const preconfiguredEntries = manifest.preconfiguredEntries[0];
        this.setData(Object.assign({}, this.data, {
            id,
            webPartData: {
                dataVersion: "1.0",
                description: preconfiguredEntries.description.default,
                id: componendId,
                instanceId: id,
                properties: preconfiguredEntries.properties,
                title: preconfiguredEntries.title.default,
            },
            webPartId: componendId,
        }));
    }
}
ClientsideWebpart.Default = {
    addedFromPersistedData: false,
    controlType: 3,
    displayMode: 2,
    emphasis: {},
    id: null,
    position: {
        controlIndex: 1,
        layoutIndex: 1,
        sectionFactor: 12,
        sectionIndex: 1,
        zoneIndex: 1,
    },
    reservedHeight: 500,
    reservedWidth: 500,
    webPartData: null,
    webPartId: null,
};


/***/ }),

/***/ 3354:
/*!******************************************************!*\
  !*** ./node_modules/@pnp/sp/clientside-pages/web.js ***!
  \******************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 7160);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _pnp_sp__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @pnp/sp */ 7881);





_webs_types_js__WEBPACK_IMPORTED_MODULE_0__._Web.prototype.getClientsideWebParts = function () {
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPCollection)(this, "GetClientSideWebParts")();
};
_webs_types_js__WEBPACK_IMPORTED_MODULE_0__._Web.prototype.addClientsidePage =
    function (pageName, title = pageName.replace(/\.[^/.]+$/, ""), layout, promotedState) {
        return (0,_types_js__WEBPACK_IMPORTED_MODULE_1__.CreateClientsidePage)(this, pageName, title, layout, promotedState);
    };
_webs_types_js__WEBPACK_IMPORTED_MODULE_0__._Web.prototype.loadClientsidePage = function (path) {
    return (0,_types_js__WEBPACK_IMPORTED_MODULE_1__.ClientsidePageFromFile)(this.getFileByServerRelativePath(path));
};
_webs_types_js__WEBPACK_IMPORTED_MODULE_0__._Web.prototype.addRepostPage = async function (details) {
    const query = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPInstance)([this, (0,_pnp_sp__WEBPACK_IMPORTED_MODULE_4__.extractWebUrl)(this.toUrl())], "_api/sitepages/pages/reposts");
    const r = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(query, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_3__.body)(details));
    return r.AbsoluteUrl;
};
// eslint-disable-next-line max-len
_webs_types_js__WEBPACK_IMPORTED_MODULE_0__._Web.prototype.addFullPageApp = async function (pageName, title = pageName.replace(/\.[^/.]+$/, ""), componentId, promotedState) {
    const parts = await this.getClientsideWebParts();
    const test = new RegExp(`{?${componentId}}?`, "i");
    const partDef = parts.find(p => test.test(p.Id));
    const part = _types_js__WEBPACK_IMPORTED_MODULE_1__.ClientsideWebpart.fromComponentDef(partDef);
    const page = await this.addClientsidePage(pageName, title, "SingleWebPartAppPage", promotedState);
    page.addSection().addColumn(12).addControl(part);
    return page;
};


/***/ }),

/***/ 914:
/*!********************************************************!*\
  !*** ./node_modules/@pnp/sp/column-defaults/folder.js ***!
  \********************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../utils/odata-url-from.js */ 905);
/* harmony import */ var _utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../utils/extract-web-url.js */ 2997);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _lists_web_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../lists/web.js */ 2519);
/* harmony import */ var _folders_types_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../folders/types.js */ 187);





_folders_types_js__WEBPACK_IMPORTED_MODULE_4__._Folder.prototype.getDefaultColumnValues = async function () {
    const folderProps = await (0,_folders_types_js__WEBPACK_IMPORTED_MODULE_4__.Folder)(this, "Properties").select("vti_x005f_listname")();
    const { ServerRelativePath: serRelPath } = await this.select("ServerRelativePath")();
    const web = (0,_webs_types_js__WEBPACK_IMPORTED_MODULE_2__.Web)([this, (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_1__.extractWebUrl)((0,_utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_0__.odataUrlFrom)(folderProps))]);
    const docLib = web.lists.getById(folderProps.vti_x005f_listname);
    // and we return the defaults associated with this folder's server relative path only
    // if you want all the defaults use list.getDefaultColumnValues()
    return (await docLib.getDefaultColumnValues()).filter(v => v.path.toLowerCase() === serRelPath.DecodedUrl.toLowerCase());
};
_folders_types_js__WEBPACK_IMPORTED_MODULE_4__._Folder.prototype.setDefaultColumnValues = async function (fieldDefaults, merge = true) {
    // we start by figuring out where we are
    const folderProps = await (0,_folders_types_js__WEBPACK_IMPORTED_MODULE_4__.Folder)(this, "Properties").select("vti_x005f_listname")();
    // now we create a web, list and batch to get some info we need
    const web = (0,_webs_types_js__WEBPACK_IMPORTED_MODULE_2__.Web)([this, (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_1__.extractWebUrl)((0,_utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_0__.odataUrlFrom)(folderProps))]);
    const docLib = web.lists.getById(folderProps.vti_x005f_listname);
    // we need the proper folder path
    const folderPath = (await this.select("ServerRelativePath")()).ServerRelativePath.DecodedUrl;
    // at this point we should have all the defaults to update
    // and we need to get all the defaults to update the entire doc
    const existingDefaults = await docLib.getDefaultColumnValues();
    // we filter all defaults to remove any associated with this folder if merge is false
    const filteredExistingDefaults = merge ? existingDefaults : existingDefaults.filter(f => f.path !== folderPath);
    // we update / add any new defaults from those passed to this method
    fieldDefaults.forEach(d => {
        const existing = filteredExistingDefaults.find(ed => ed.name === d.name && ed.path === folderPath);
        if (existing) {
            existing.value = d.value;
        }
        else {
            filteredExistingDefaults.push({
                name: d.name,
                path: folderPath,
                value: d.value,
            });
        }
    });
    // after this operation filteredExistingDefaults should contain all the value we want to write to the file
    await docLib.setDefaultColumnValues(filteredExistingDefaults);
};
_folders_types_js__WEBPACK_IMPORTED_MODULE_4__._Folder.prototype.clearDefaultColumnValues = async function () {
    await this.setDefaultColumnValues([], false);
};


/***/ }),

/***/ 4763:
/*!*******************************************************!*\
  !*** ./node_modules/@pnp/sp/column-defaults/index.js ***!
  \*******************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _list_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./list.js */ 1392);
/* harmony import */ var _folder_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./folder.js */ 914);




/***/ }),

/***/ 1392:
/*!******************************************************!*\
  !*** ./node_modules/@pnp/sp/column-defaults/list.js ***!
  \******************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _lists_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types.js */ 9706);
/* harmony import */ var _folders_types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../folders/types.js */ 187);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _presets_all_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../presets/all.js */ 6935);
/* harmony import */ var _utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../utils/encode-path-str.js */ 4729);







(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List, "rootFolder", _folders_types_js__WEBPACK_IMPORTED_MODULE_2__.Folder);
_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List.prototype.getDefaultColumnValues = async function () {
    const pathPart = await this.rootFolder.select("ServerRelativePath")();
    const webUrl = await this.select("ParentWeb/Url").expand("ParentWeb")();
    const path = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_3__.combine)("/", pathPart.ServerRelativePath.DecodedUrl, "Forms/client_LocationBasedDefaults.html");
    const baseFilePath = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_3__.combine)(webUrl.ParentWeb.Url, `_api/web/getFileByServerRelativePath(decodedUrl='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__.encodePath)(path)}')`);
    let xml = "";
    try {
        // we are reading the contents of the file
        xml = await (0,_folders_types_js__WEBPACK_IMPORTED_MODULE_2__.Folder)([this, baseFilePath], "$value").using((0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.TextParse)())();
    }
    catch (e) {
        // if this call fails we assume it is because the file is 404
        if (e && e.status && e.status === 404) {
            // return an empty array
            return [];
        }
        throw e;
    }
    // get all the tags from the xml
    const matches = xml.match(/<a.*?<\/a>/ig);
    const tags = matches === null ? [] : matches.map(t => t.trim());
    // now we need to turn these tags of form into objects
    // <a href="/sites/dev/My%20Title"><DefaultValue FieldName="TextField">Test</DefaultValue></a>
    return tags.reduce((defVals, t) => {
        const m = /<a href="(.*?)">/ig.exec(t);
        // if things worked out captures are:
        // 0: whole string
        // 1: ENCODED server relative path
        if (m.length < 1) {
            // this indicates an error somewhere, but we have no way to meaningfully recover
            // perhaps the way the tags are stored has changed on the server? Check that first.
            this.log(`Could not parse default column value from '${t}'`, 2);
            return null;
        }
        // return the parsed out values
        const subMatches = t.match(/<DefaultValue.*?<\/DefaultValue>/ig);
        const subTags = subMatches === null ? [] : subMatches.map(st => st.trim());
        subTags.map(st => {
            const sm = /<DefaultValue FieldName="(.*?)">(.*?)<\/DefaultValue>/ig.exec(st);
            // if things worked out captures are:
            // 0: whole string
            // 1: Field internal name
            // 2: Default value as string
            if (sm.length < 1) {
                this.log(`Could not parse default column value from '${st}'`, 2);
            }
            else {
                defVals.push({
                    name: sm[1],
                    path: decodeURIComponent(m[1]),
                    value: sm[2],
                });
            }
        });
        return defVals;
    }, []).filter(v => v !== null);
};
_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List.prototype.setDefaultColumnValues = async function (defaults) {
    // we need the field types from the list to map the values
    // eslint-disable-next-line max-len
    const fieldDefs = await (0,_presets_all_js__WEBPACK_IMPORTED_MODULE_5__.SPCollection)(this, "fields").select("InternalName", "TypeAsString").filter("Hidden ne true")();
    // group field defaults by path
    const defaultsByPath = {};
    for (let i = 0; i < defaults.length; i++) {
        if (defaultsByPath[defaults[i].path] == null) {
            defaultsByPath[defaults[i].path] = [defaults[i]];
        }
        else {
            defaultsByPath[defaults[i].path].push(defaults[i]);
        }
    }
    const paths = Object.getOwnPropertyNames(defaultsByPath);
    const pathDefaults = [];
    // For each path, group field defaults
    for (let j = 0; j < paths.length; j++) {
        // map the values into the right format and produce our xml elements
        const pathFields = defaultsByPath[paths[j]];
        const tags = pathFields.map(fieldDefault => {
            const index = fieldDefs.findIndex(fd => fd.InternalName === fieldDefault.name);
            if (index < 0) {
                throw Error(`Field '${fieldDefault.name}' does not exist in the list. Please check the internal field name. Failed to set defaults.`);
            }
            const fieldDef = fieldDefs[index];
            let value = "";
            switch (fieldDef.TypeAsString) {
                case "Boolean":
                case "Currency":
                case "Text":
                case "DateTime":
                case "Number":
                case "Choice":
                case "User":
                    if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_3__.isArray)(fieldDefault.value)) {
                        throw Error(`The type '${fieldDef.TypeAsString}' does not support multiple values.`);
                    }
                    value = `${fieldDefault.value}`;
                    break;
                case "MultiChoice":
                    if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_3__.isArray)(fieldDefault.value)) {
                        value = fieldDefault.value.map(v => `${v}`).join(";");
                    }
                    else {
                        value = `${fieldDefault.value}`;
                    }
                    break;
                case "UserMulti":
                    if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_3__.isArray)(fieldDefault.value)) {
                        value = fieldDefault.value.map(v => `${v}`).join(";#");
                    }
                    else {
                        value = `${fieldDefault.value}`;
                    }
                    break;
                case "Taxonomy":
                case "TaxonomyFieldType":
                    if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_3__.isArray)(fieldDefault.value)) {
                        throw Error(`The type '${fieldDef.TypeAsString}' does not support multiple values.`);
                    }
                    else {
                        value = `${fieldDefault.value.wssId};#${fieldDefault.value.termName}|${fieldDefault.value.termId}`;
                    }
                    break;
                case "TaxonomyMulti":
                case "TaxonomyFieldTypeMulti":
                    if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_3__.isArray)(fieldDefault.value)) {
                        value = fieldDefault.value.map(v => `${v.wssId};#${v.termName}|${v.termId}`).join(";#");
                    }
                    else {
                        value = [fieldDefault.value].map(v => `${v.wssId};#${v.termName}|${v.termId}`).join(";#");
                    }
                    break;
            }
            return `<DefaultValue FieldName="${fieldDefault.name}">${value}</DefaultValue>`;
        });
        const href = pathFields[0].path.replace(/ /gi, "%20");
        const pathDefault = `<a href="${href}">${tags.join("")}</a>`;
        pathDefaults.push(pathDefault);
    }
    // builds update to defaults
    const xml = `<MetadataDefaults>${pathDefaults.join("")}</MetadataDefaults>`;
    const pathPart = await this.rootFolder.select("ServerRelativePath")();
    const webUrl = await this.select("ParentWeb/Url").expand("ParentWeb")();
    const path = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_3__.combine)("/", pathPart.ServerRelativePath.DecodedUrl, "Forms");
    const baseFilePath = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_3__.combine)(webUrl.ParentWeb.Url, "_api/web", `getFolderByServerRelativePath(decodedUrl='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__.encodePath)(path)}')`, "files");
    await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_4__.spPost)((0,_folders_types_js__WEBPACK_IMPORTED_MODULE_2__.Folder)([this, baseFilePath], "add(overwrite=true,url='client_LocationBasedDefaults.html')"), { body: xml });
    // finally we need to ensure this list has the right event receiver added
    const existingReceivers = await this.eventReceivers.filter("ReceiverName eq 'LocationBasedMetadataDefaultsReceiver ItemAdded'").select("ReceiverId")();
    if (existingReceivers.length < 1) {
        await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_4__.spPost)((0,_lists_types_js__WEBPACK_IMPORTED_MODULE_1__.List)(this.eventReceivers, "add"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({
            eventReceiverCreationInformation: {
                EventType: 10001,
                ReceiverAssembly: "Microsoft.Office.DocumentManagement, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c",
                ReceiverClass: "Microsoft.Office.DocumentManagement.LocationBasedMetadataDefaultsReceiver",
                ReceiverName: "LocationBasedMetadataDefaultsReceiver ItemAdded",
                SequenceNumber: 1000,
                Synchronization: 1,
            },
        }));
    }
};


/***/ }),

/***/ 5429:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/comments/index.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _item_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./item.js */ 6823);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 4569);




/***/ }),

/***/ 6823:
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/comments/item.js ***!
  \***********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _items_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../items/types.js */ 132);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 4569);
/* harmony import */ var _utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../utils/extract-web-url.js */ 2997);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../spqueryable.js */ 2678);






(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_items_types_js__WEBPACK_IMPORTED_MODULE_1__._Item, "comments", _types_js__WEBPACK_IMPORTED_MODULE_2__.Comments);
_items_types_js__WEBPACK_IMPORTED_MODULE_1__._Item.prototype.getLikedBy = function () {
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_5__.spPost)((0,_items_types_js__WEBPACK_IMPORTED_MODULE_1__.Item)(this, "likedBy"));
};
_items_types_js__WEBPACK_IMPORTED_MODULE_1__._Item.prototype.like = async function () {
    const itemInfo = await this.getParentInfos();
    const baseUrl = (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_3__.extractWebUrl)(this.toUrl());
    const reputationUrl = "_api/Microsoft.Office.Server.ReputationModel.Reputation.SetLike(listID=@a1,itemID=@a2,like=@a3)";
    const likeUrl = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_4__.combine)(baseUrl, reputationUrl) + `?@a1='{${itemInfo.ParentList.Id}}'&@a2='${itemInfo.Item.Id}'&@a3=true`;
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_5__.spPost)((0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_5__.SPQueryable)([this, likeUrl]));
};
_items_types_js__WEBPACK_IMPORTED_MODULE_1__._Item.prototype.unlike = async function () {
    const itemInfo = await this.getParentInfos();
    const baseUrl = (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_3__.extractWebUrl)(this.toUrl());
    const reputationUrl = "_api/Microsoft.Office.Server.ReputationModel.Reputation.SetLike(listID=@a1,itemID=@a2,like=@a3)";
    const likeUrl = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_4__.combine)(baseUrl, reputationUrl) + `?@a1='{${itemInfo.ParentList.Id}}'&@a2='${itemInfo.Item.Id}'&@a3=false`;
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_5__.spPost)((0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_5__.SPQueryable)([this, likeUrl]));
};
_items_types_js__WEBPACK_IMPORTED_MODULE_1__._Item.prototype.getLikedByInformation = function () {
    return (0,_items_types_js__WEBPACK_IMPORTED_MODULE_1__.Item)(this, "likedByInformation").expand("likedby")();
};
_items_types_js__WEBPACK_IMPORTED_MODULE_1__._Item.prototype.rate = async function (value) {
    const itemInfo = await this.getParentInfos();
    const baseUrl = (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_3__.extractWebUrl)(this.toUrl());
    const reputationUrl = "_api/Microsoft.Office.Server.ReputationModel.Reputation.SetRating(listID=@a1,itemID=@a2,rating=@a3)";
    const rateUrl = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_4__.combine)(baseUrl, reputationUrl) + `?@a1='{${itemInfo.ParentList.Id}}'&@a2='${itemInfo.Item.Id}'&@a3=${value}`;
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_5__.spPost)((0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_5__.SPQueryable)([this, rateUrl]));
};


/***/ }),

/***/ 4569:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/comments/types.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Comments: () => (/* binding */ Comments)
/* harmony export */ });
/* unused harmony exports _Comments, _Comment, Comment, _Replies, Replies */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../decorators.js */ 6540);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../utils/odata-url-from.js */ 905);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/queryable */ 6844);





let _Comments = class _Comments extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPCollection {
    /**
     * Adds a new comment to this collection
     *
     * @param info Comment information to add
     */
    async add(info) {
        if (typeof info === "string") {
            info = { text: info };
        }
        const d = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(Comments(this, null), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_2__.body)(info));
        return Object.assign(this.getById(d.id), d);
    }
    /**
     * Gets a comment by id
     *
     * @param id Id of the comment to load
     */
    getById(id) {
        return Comment(this).concat(`(${id})`);
    }
    /**
     * Deletes all the comments in this collection
     */
    clear() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(Comments(this, "DeleteAll"));
    }
};
_Comments = (0,tslib__WEBPACK_IMPORTED_MODULE_3__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_4__.defaultPath)("comments")
], _Comments);

const Comments = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_Comments);
class _Comment extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPInstance {
    /**
     * A comment's replies
     */
    get replies() {
        return Replies(this);
    }
    /**
     * Likes the comment as the current user
     */
    like() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(Comment(this, "Like"));
    }
    /**
     * Unlikes the comment as the current user
     */
    unlike() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(Comment(this, "Unlike"));
    }
    /**
     * Deletes this comment
     */
    delete() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spDelete)(this);
    }
}
const Comment = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_Comment);
let _Replies = class _Replies extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPCollection {
    /**
     * Adds a new reply to this collection
     *
     * @param info Comment information to add
     */
    async add(info) {
        if (typeof info === "string") {
            info = { text: info };
        }
        const d = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(Replies(this, null), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_2__.body)(info));
        return Object.assign(Comment([this, (0,_utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_1__.odataUrlFrom)(d)]), d);
    }
};
_Replies = (0,tslib__WEBPACK_IMPORTED_MODULE_3__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_4__.defaultPath)("replies")
], _Replies);

const Replies = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_Replies);


/***/ }),

/***/ 6959:
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/sp/content-types/index.js ***!
  \*****************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _web_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web.js */ 5068);
/* harmony import */ var _item_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./item.js */ 8775);
/* harmony import */ var _list_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./list.js */ 1598);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./types.js */ 8203);






/***/ }),

/***/ 8775:
/*!****************************************************!*\
  !*** ./node_modules/@pnp/sp/content-types/item.js ***!
  \****************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _items_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../items/types.js */ 132);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 8203);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_items_types_js__WEBPACK_IMPORTED_MODULE_1__._Item, "contentType", _types_js__WEBPACK_IMPORTED_MODULE_2__.ContentType);


/***/ }),

/***/ 1598:
/*!****************************************************!*\
  !*** ./node_modules/@pnp/sp/content-types/list.js ***!
  \****************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _lists_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types.js */ 9706);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 8203);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List, "contentTypes", _types_js__WEBPACK_IMPORTED_MODULE_2__.ContentTypes);


/***/ }),

/***/ 8203:
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/sp/content-types/types.js ***!
  \*****************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   ContentType: () => (/* binding */ ContentType),
/* harmony export */   ContentTypes: () => (/* binding */ ContentTypes),
/* harmony export */   _ContentType: () => (/* binding */ _ContentType)
/* harmony export */ });
/* unused harmony exports _ContentTypes, _FieldLinks, FieldLinks, _FieldLink, FieldLink */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../decorators.js */ 6540);




let _ContentTypes = class _ContentTypes extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPCollection {
    /**
     * Adds an existing contenttype to a content type collection
     *
     * @param contentTypeId in the following format, for example: 0x010102
     */
    async addAvailableContentType(contentTypeId) {
        const data = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(ContentTypes(this, "addAvailableContentType"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({ "contentTypeId": contentTypeId }));
        return {
            contentType: this.getById(data.id),
            data: data,
        };
    }
    /**
     * Gets a ContentType by content type id
     * @param id The id of the content type to get, in the following format, for example: 0x010102
     */
    getById(id) {
        return ContentType(this).concat(`('${id}')`);
    }
    /**
     * Adds a new content type to the collection
     *
     * @param id The desired content type id for the new content type (also determines the parent
     *   content type)
     * @param name The name of the content type
     * @param description The description of the content type
     * @param group The group in which to add the content type
     * @param additionalSettings Any additional settings to provide when creating the content type
     *
     */
    async add(id, name, description = "", group = "Custom Content Types", additionalSettings = {}) {
        const postBody = (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({
            Description: description,
            Group: group,
            Id: { StringValue: id },
            Name: name,
            ...additionalSettings,
        });
        const data = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(this, postBody);
        return { contentType: this.getById(data.id), data };
    }
};
_ContentTypes = (0,tslib__WEBPACK_IMPORTED_MODULE_2__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_3__.defaultPath)("contenttypes")
], _ContentTypes);

const ContentTypes = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spInvokableFactory)(_ContentTypes);
class _ContentType extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPInstance {
    constructor() {
        super(...arguments);
        this.delete = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.deleteable)();
    }
    /**
     * Updates this list instance with the supplied properties
     *
     * @param properties A plain object hash of values to update for the web
     */
    async update(properties) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPostMerge)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)(properties));
    }
    /**
     * Gets the column (also known as field) references in the content type.
     */
    get fieldLinks() {
        return FieldLinks(this);
    }
    /**
     * Gets a value that specifies the collection of fields for the content type.
     */
    get fields() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.SPCollection)(this, "fields");
    }
    /**
     * Gets the parent content type of the content type.
     */
    get parent() {
        return ContentType(this, "parent");
    }
    /**
     * Gets a value that specifies the collection of workflow associations for the content type.
     */
    get workflowAssociations() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.SPCollection)(this, "workflowAssociations");
    }
}
const ContentType = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spInvokableFactory)(_ContentType);
let _FieldLinks = class _FieldLinks extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPCollection {
    /**
    *  Gets a FieldLink by GUID id
    *
    * @param id The GUID id of the field link
    */
    getById(id) {
        return FieldLink(this).concat(`(guid'${id}')`);
    }
};
_FieldLinks = (0,tslib__WEBPACK_IMPORTED_MODULE_2__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_3__.defaultPath)("fieldlinks")
], _FieldLinks);

const FieldLinks = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spInvokableFactory)(_FieldLinks);
class _FieldLink extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPInstance {
}
const FieldLink = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spInvokableFactory)(_FieldLink);


/***/ }),

/***/ 5068:
/*!***************************************************!*\
  !*** ./node_modules/@pnp/sp/content-types/web.js ***!
  \***************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 8203);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web, "contentTypes", _types_js__WEBPACK_IMPORTED_MODULE_2__.ContentTypes);


/***/ }),

/***/ 1734:
/*!****************************************************!*\
  !*** ./node_modules/@pnp/sp/context-info/index.js ***!
  \****************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../utils/extract-web-url.js */ 2997);



_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPQueryable.prototype.getContextInfo = async function (path = this.parentUrl) {
    const data = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)((0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.SPQueryable)([this, (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_2__.extractWebUrl)(path)], "_api/contextinfo"));
    if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(data, "GetContextWebInformation")) {
        const info = data.GetContextWebInformation;
        info.SupportedSchemaVersions = info.SupportedSchemaVersions.results;
        return info;
    }
    else {
        return data;
    }
};


/***/ }),

/***/ 6540:
/*!********************************************!*\
  !*** ./node_modules/@pnp/sp/decorators.js ***!
  \********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   defaultPath: () => (/* binding */ defaultPath)
/* harmony export */ });
/**
 * Decorator used to specify the default path for SPQueryable objects
 *
 * @param path
 */
function defaultPath(path) {
    // eslint-disable-next-line @typescript-eslint/ban-types
    return function (target) {
        return class extends target {
            constructor(...args) {
                super(args[0], args.length > 1 && args[1] !== undefined ? args[1] : path);
            }
        };
    };
}


/***/ }),

/***/ 9382:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/features/index.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _site_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./site.js */ 381);
/* harmony import */ var _web_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./web.js */ 9940);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 845);





/***/ }),

/***/ 381:
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/features/site.js ***!
  \***********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _sites_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sites/types.js */ 1114);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 845);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_sites_types_js__WEBPACK_IMPORTED_MODULE_1__._Site, "features", _types_js__WEBPACK_IMPORTED_MODULE_2__.Features);


/***/ }),

/***/ 845:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/features/types.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Features: () => (/* binding */ Features)
/* harmony export */ });
/* unused harmony exports _Features, _Feature, Feature */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../decorators.js */ 6540);




let _Features = class _Features extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPCollection {
    /**
     * Adds (activates) the specified feature
     *
     * @param id The Id of the feature (GUID)
     * @param force If true the feature activation will be forced
     */
    async add(featureId, force = false) {
        const data = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(Features(this, "add"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({
            featdefScope: 0,
            featureId,
            force,
        }));
        return {
            data: data,
            feature: this.getById(featureId),
        };
    }
    /**
     * Gets a feature from the collection with the specified guid
     *
     * @param id The Id of the feature (GUID)
     */
    getById(id) {
        return Feature(this).concat(`('${id}')`);
    }
    /**
     * Removes (deactivates) a feature from the collection
     *
     * @param id The Id of the feature (GUID)
     * @param force If true the feature deactivation will be forced
     */
    remove(featureId, force = false) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(Features(this, "remove"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({
            featureId,
            force,
        }));
    }
};
_Features = (0,tslib__WEBPACK_IMPORTED_MODULE_2__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_3__.defaultPath)("features")
], _Features);

const Features = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spInvokableFactory)(_Features);
class _Feature extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPInstance {
}
const Feature = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spInvokableFactory)(_Feature);


/***/ }),

/***/ 9940:
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/features/web.js ***!
  \**********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 845);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web, "features", _types_js__WEBPACK_IMPORTED_MODULE_2__.Features);


/***/ }),

/***/ 6553:
/*!************************************!*\
  !*** ./node_modules/@pnp/sp/fi.js ***!
  \************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   SPFI: () => (/* binding */ SPFI),
/* harmony export */   spfi: () => (/* binding */ spfi)
/* harmony export */ });
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./spqueryable.js */ 2678);

class SPFI {
    /**
     * Creates a new instance of the SPFI class
     *
     * @param root Establishes a root url/configuration
     */
    constructor(root = "") {
        this._root = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.SPQueryable)(root);
    }
    /**
     * Applies one or more behaviors which will be inherited by all instances chained from this root
     *
     */
    using(...behaviors) {
        this._root.using(...behaviors);
        return this;
    }
    /**
     * Used by extending classes to create new objects directly from the root
     *
     * @param factory The factory for the type of object to create
     * @returns A configured instance of that object
     */
    create(factory, path) {
        return factory(this._root, path);
    }
}
function spfi(root = "") {
    if (typeof root === "object" && !Reflect.has(root, "length")) {
        root = root._root;
    }
    return new SPFI(root);
}


/***/ }),

/***/ 3032:
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/fields/index.js ***!
  \**********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _web_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web.js */ 7580);
/* harmony import */ var _list_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./list.js */ 5213);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 5472);





/***/ }),

/***/ 5213:
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/fields/list.js ***!
  \*********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _lists_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types.js */ 9706);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 5472);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List, "fields", _types_js__WEBPACK_IMPORTED_MODULE_2__.Fields);


/***/ }),

/***/ 5472:
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/fields/types.js ***!
  \**********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Fields: () => (/* binding */ Fields),
/* harmony export */   _Field: () => (/* binding */ _Field)
/* harmony export */ });
/* unused harmony exports _Fields, Field, FieldTypes, DateTimeFieldFormatType, DateTimeFieldFriendlyFormatType, AddFieldOptions, CalendarType, UrlFieldFormatType, FieldUserSelectionMode, ChoiceFieldFormatType */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../decorators.js */ 6540);
/* harmony import */ var _utils_metadata_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../utils/metadata.js */ 3999);





let _Fields = class _Fields extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPCollection {
    /**
     * Creates a field based on the specified schema
     *
     * @param xml A string or XmlSchemaFieldCreationInformation instance descrbing the field to create
     */
    async createFieldAsXml(xml) {
        if (typeof xml === "string") {
            xml = { SchemaXml: xml };
        }
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(Fields(this, "createfieldasxml"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({ parameters: xml }));
    }
    /**
     * Gets a field from the collection by id
     *
     * @param id The Id of the list
     */
    getById(id) {
        return Field(this).concat(`('${id}')`);
    }
    /**
     * Gets a field from the collection by title
     *
     * @param title The case-sensitive title of the field
     */
    getByTitle(title) {
        return Field(this, `getByTitle('${title}')`);
    }
    /**
     * Gets a field from the collection by using internal name or title
     *
     * @param name The case-sensitive internal name or title of the field
     */
    getByInternalNameOrTitle(name) {
        return Field(this, `getByInternalNameOrTitle('${name}')`);
    }
    /**
     * Adds a new field to the collection
     *
     * @param title The new field's title
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    async add(title, fieldTypeKind, properties) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(Fields(this, null), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)(Object.assign((0,_utils_metadata_js__WEBPACK_IMPORTED_MODULE_2__.metadata)(mapFieldTypeEnumToString(fieldTypeKind)), {
            Title: title,
            FieldTypeKind: fieldTypeKind,
            ...properties,
        }), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.headers)({
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
        })));
    }
    /**
     * Adds a new field to the collection
     *
     * @param title The new field's title
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    async addField(title, fieldTypeKind, properties) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(Fields(this, "AddField"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({
            parameters: {
                Title: title,
                FieldTypeKind: fieldTypeKind,
                ...properties,
            },
        }));
    }
    /**
     * Adds a new SP.FieldText to the collection
     *
     * @param title The field title
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    addText(title, properties) {
        return this.add(title, 2, {
            MaxLength: 255,
            ...properties,
        });
    }
    /**
     * Adds a new SP.FieldCalculated to the collection
     *
     * @param title The field title.
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    addCalculated(title, properties) {
        return this.add(title, 17, {
            OutputType: FieldTypes.Text,
            ...properties,
        });
    }
    /**
     * Adds a new SP.FieldDateTime to the collection
     *
     * @param title The field title
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    addDateTime(title, properties) {
        return this.add(title, 4, {
            DateTimeCalendarType: CalendarType.Gregorian,
            DisplayFormat: DateTimeFieldFormatType.DateOnly,
            FriendlyDisplayFormat: DateTimeFieldFriendlyFormatType.Unspecified,
            ...properties,
        });
    }
    /**
     * Adds a new SP.FieldNumber to the collection
     *
     * @param title The field title
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    addNumber(title, properties) {
        return this.add(title, 9, properties);
    }
    /**
     * Adds a new SP.FieldCurrency to the collection
     *
     * @param title The field title
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    addCurrency(title, properties) {
        return this.add(title, 10, {
            CurrencyLocaleId: 1033,
            ...properties,
        });
    }
    /**
     * Adds a new SP.FieldMultiLineText to the collection
     *
     * @param title The field title
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     *
     */
    addMultilineText(title, properties) {
        return this.add(title, 3, {
            AllowHyperlink: true,
            AppendOnly: false,
            NumberOfLines: 6,
            RestrictedMode: false,
            RichText: true,
            ...properties,
        });
    }
    /**
     * Adds a new SP.FieldUrl to the collection
     *
     * @param title The field title
     */
    addUrl(title, properties) {
        return this.add(title, 11, {
            DisplayFormat: UrlFieldFormatType.Hyperlink,
            ...properties,
        });
    }
    /** Adds a user field to the colleciton
     *
     * @param title The new field's title
     * @param properties
     */
    addUser(title, properties) {
        return this.add(title, 20, {
            SelectionMode: FieldUserSelectionMode.PeopleAndGroups,
            ...properties,
        });
    }
    /**
     * Adds a SP.FieldLookup to the collection
     *
     * @param title The new field's title
     * @param properties Set of additional properties to set on the new field
     */
    async addLookup(title, properties) {
        return this.addField(title, 7, properties);
    }
    /**
     * Adds a new SP.FieldChoice to the collection
     *
     * @param title The field title.
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    addChoice(title, properties) {
        const props = {
            ...properties,
            Choices: {
                results: properties.Choices,
            },
        };
        return this.add(title, 6, props);
    }
    /**
     * Adds a new SP.FieldMultiChoice to the collection
     *
     * @param title The field title.
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    addMultiChoice(title, properties) {
        const props = {
            ...properties,
            Choices: {
                results: properties.Choices,
            },
        };
        return this.add(title, 15, props);
    }
    /**
   * Adds a new SP.FieldBoolean to the collection
   *
   * @param title The field title.
   * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
   */
    addBoolean(title, properties) {
        return this.add(title, 8, properties);
    }
    /**
  * Creates a secondary (dependent) lookup field, based on the Id of the primary lookup field.
  *
  * @param displayName The display name of the new field.
  * @param primaryLookupFieldId The guid of the primary Lookup Field.
  * @param showField Which field to show from the lookup list.
  */
    async addDependentLookupField(displayName, primaryLookupFieldId, showField) {
        const path = `adddependentlookupfield(displayName='${displayName}', primarylookupfieldid='${primaryLookupFieldId}', showfield='${showField}')`;
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(Fields(this, path));
    }
    /**
   * Adds a new SP.FieldLocation to the collection
   *
   * @param title The field title.
   * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
   */
    addLocation(title, properties) {
        return this.add(title, 33, properties);
    }
    /**
     * Adds a new SP.FieldLocation to the collection
     *
     * @param title The field title.
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    addImageField(title, properties) {
        return this.add(title, 34, properties);
    }
};
_Fields = (0,tslib__WEBPACK_IMPORTED_MODULE_3__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_4__.defaultPath)("fields")
], _Fields);

const Fields = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spInvokableFactory)(_Fields);
class _Field extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPInstance {
    constructor() {
        super(...arguments);
        this.delete = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.deleteable)();
    }
    /**
   * Updates this field instance with the supplied properties
   *
   * @param properties A plain object hash of values to update for the list
   * @param fieldType The type value such as SP.FieldLookup. Optional, looked up from the field if not provided
   */
    async update(properties, fieldType) {
        if (typeof fieldType === "undefined" || fieldType === null) {
            const info = await Field(this).select("FieldTypeKind")();
            fieldType = info["odata.type"];
        }
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPostMerge)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)(properties));
    }
    /**
   * Sets the value of the ShowInDisplayForm property for this field.
   */
    setShowInDisplayForm(show) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(Field(this, `setshowindisplayform(${show})`));
    }
    /**
   * Sets the value of the ShowInEditForm property for this field.
   */
    setShowInEditForm(show) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(Field(this, `setshowineditform(${show})`));
    }
    /**
   * Sets the value of the ShowInNewForm property for this field.
   */
    setShowInNewForm(show) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(Field(this, `setshowinnewform(${show})`));
    }
}
const Field = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spInvokableFactory)(_Field);
/**
 * Specifies the type of the field.
 */
var FieldTypes;
(function (FieldTypes) {
    FieldTypes[FieldTypes["Invalid"] = 0] = "Invalid";
    FieldTypes[FieldTypes["Integer"] = 1] = "Integer";
    FieldTypes[FieldTypes["Text"] = 2] = "Text";
    FieldTypes[FieldTypes["Note"] = 3] = "Note";
    FieldTypes[FieldTypes["DateTime"] = 4] = "DateTime";
    FieldTypes[FieldTypes["Counter"] = 5] = "Counter";
    FieldTypes[FieldTypes["Choice"] = 6] = "Choice";
    FieldTypes[FieldTypes["Lookup"] = 7] = "Lookup";
    FieldTypes[FieldTypes["Boolean"] = 8] = "Boolean";
    FieldTypes[FieldTypes["Number"] = 9] = "Number";
    FieldTypes[FieldTypes["Currency"] = 10] = "Currency";
    FieldTypes[FieldTypes["URL"] = 11] = "URL";
    FieldTypes[FieldTypes["Computed"] = 12] = "Computed";
    FieldTypes[FieldTypes["Threading"] = 13] = "Threading";
    FieldTypes[FieldTypes["Guid"] = 14] = "Guid";
    FieldTypes[FieldTypes["MultiChoice"] = 15] = "MultiChoice";
    FieldTypes[FieldTypes["GridChoice"] = 16] = "GridChoice";
    FieldTypes[FieldTypes["Calculated"] = 17] = "Calculated";
    FieldTypes[FieldTypes["File"] = 18] = "File";
    FieldTypes[FieldTypes["Attachments"] = 19] = "Attachments";
    FieldTypes[FieldTypes["User"] = 20] = "User";
    FieldTypes[FieldTypes["Recurrence"] = 21] = "Recurrence";
    FieldTypes[FieldTypes["CrossProjectLink"] = 22] = "CrossProjectLink";
    FieldTypes[FieldTypes["ModStat"] = 23] = "ModStat";
    FieldTypes[FieldTypes["Error"] = 24] = "Error";
    FieldTypes[FieldTypes["ContentTypeId"] = 25] = "ContentTypeId";
    FieldTypes[FieldTypes["PageSeparator"] = 26] = "PageSeparator";
    FieldTypes[FieldTypes["ThreadIndex"] = 27] = "ThreadIndex";
    FieldTypes[FieldTypes["WorkflowStatus"] = 28] = "WorkflowStatus";
    FieldTypes[FieldTypes["AllDayEvent"] = 29] = "AllDayEvent";
    FieldTypes[FieldTypes["WorkflowEventType"] = 30] = "WorkflowEventType";
    FieldTypes[FieldTypes["Location"] = 33] = "Location";
    FieldTypes[FieldTypes["Image"] = 34] = "Image";
})(FieldTypes || (FieldTypes = {}));
const FieldTypeClassMapping = {
    [FieldTypes.Calculated]: "SP.FieldCalculated",
    [FieldTypes.Choice]: "SP.FieldChoice",
    [FieldTypes.Computed]: "SP.FieldComputed",
    [FieldTypes.Currency]: "SP.FieldCurrency",
    [FieldTypes.DateTime]: "SP.FieldDateTime",
    [FieldTypes.GridChoice]: "SP.FieldRatingScale",
    [FieldTypes.Guid]: "SP.FieldGuid",
    [FieldTypes.Image]: "SP.FieldMultiLineText",
    [FieldTypes.Integer]: "SP.FieldNumber",
    [FieldTypes.Location]: "SP.FieldLocation",
    [FieldTypes.Lookup]: "SP.FieldLookup",
    [FieldTypes.ModStat]: "SP.FieldChoice",
    [FieldTypes.MultiChoice]: "SP.FieldMultiChoice",
    [FieldTypes.Note]: "SP.FieldMultiLineText",
    [FieldTypes.Number]: "SP.FieldNumber",
    [FieldTypes.Text]: "SP.FieldText",
    [FieldTypes.URL]: "SP.FieldUrl",
    [FieldTypes.User]: "SP.FieldUser",
    [FieldTypes.WorkflowStatus]: "SP.FieldChoice",
    [FieldTypes.WorkflowEventType]: "SP.FieldNumber",
};
function mapFieldTypeEnumToString(enumValue) {
    var _a;
    return (_a = FieldTypeClassMapping[enumValue]) !== null && _a !== void 0 ? _a : "SP.Field";
}
var DateTimeFieldFormatType;
(function (DateTimeFieldFormatType) {
    DateTimeFieldFormatType[DateTimeFieldFormatType["DateOnly"] = 0] = "DateOnly";
    DateTimeFieldFormatType[DateTimeFieldFormatType["DateTime"] = 1] = "DateTime";
})(DateTimeFieldFormatType || (DateTimeFieldFormatType = {}));
var DateTimeFieldFriendlyFormatType;
(function (DateTimeFieldFriendlyFormatType) {
    DateTimeFieldFriendlyFormatType[DateTimeFieldFriendlyFormatType["Unspecified"] = 0] = "Unspecified";
    DateTimeFieldFriendlyFormatType[DateTimeFieldFriendlyFormatType["Disabled"] = 1] = "Disabled";
    DateTimeFieldFriendlyFormatType[DateTimeFieldFriendlyFormatType["Relative"] = 2] = "Relative";
})(DateTimeFieldFriendlyFormatType || (DateTimeFieldFriendlyFormatType = {}));
/**
 * Specifies the control settings while adding a field.
 */
var AddFieldOptions;
(function (AddFieldOptions) {
    /**
   *  Specify that a new field added to the list must also be added to the default content type in the site collection
   */
    AddFieldOptions[AddFieldOptions["DefaultValue"] = 0] = "DefaultValue";
    /**
   * Specify that a new field added to the list must also be added to the default content type in the site collection.
   */
    AddFieldOptions[AddFieldOptions["AddToDefaultContentType"] = 1] = "AddToDefaultContentType";
    /**
   * Specify that a new field must not be added to any other content type
   */
    AddFieldOptions[AddFieldOptions["AddToNoContentType"] = 2] = "AddToNoContentType";
    /**
   *  Specify that a new field that is added to the specified list must also be added to all content types in the site collection
   */
    AddFieldOptions[AddFieldOptions["AddToAllContentTypes"] = 4] = "AddToAllContentTypes";
    /**
   * Specify adding an internal field name hint for the purpose of avoiding possible database locking or field renaming operations
   */
    AddFieldOptions[AddFieldOptions["AddFieldInternalNameHint"] = 8] = "AddFieldInternalNameHint";
    /**
   * Specify that a new field that is added to the specified list must also be added to the default list view
   */
    AddFieldOptions[AddFieldOptions["AddFieldToDefaultView"] = 16] = "AddFieldToDefaultView";
    /**
   * Specify to confirm that no other field has the same display name
   */
    AddFieldOptions[AddFieldOptions["AddFieldCheckDisplayName"] = 32] = "AddFieldCheckDisplayName";
})(AddFieldOptions || (AddFieldOptions = {}));
var CalendarType;
(function (CalendarType) {
    CalendarType[CalendarType["Gregorian"] = 1] = "Gregorian";
    CalendarType[CalendarType["Japan"] = 3] = "Japan";
    CalendarType[CalendarType["Taiwan"] = 4] = "Taiwan";
    CalendarType[CalendarType["Korea"] = 5] = "Korea";
    CalendarType[CalendarType["Hijri"] = 6] = "Hijri";
    CalendarType[CalendarType["Thai"] = 7] = "Thai";
    CalendarType[CalendarType["Hebrew"] = 8] = "Hebrew";
    CalendarType[CalendarType["GregorianMEFrench"] = 9] = "GregorianMEFrench";
    CalendarType[CalendarType["GregorianArabic"] = 10] = "GregorianArabic";
    CalendarType[CalendarType["GregorianXLITEnglish"] = 11] = "GregorianXLITEnglish";
    CalendarType[CalendarType["GregorianXLITFrench"] = 12] = "GregorianXLITFrench";
    CalendarType[CalendarType["KoreaJapanLunar"] = 14] = "KoreaJapanLunar";
    CalendarType[CalendarType["ChineseLunar"] = 15] = "ChineseLunar";
    CalendarType[CalendarType["SakaEra"] = 16] = "SakaEra";
    CalendarType[CalendarType["UmAlQura"] = 23] = "UmAlQura";
})(CalendarType || (CalendarType = {}));
var UrlFieldFormatType;
(function (UrlFieldFormatType) {
    UrlFieldFormatType[UrlFieldFormatType["Hyperlink"] = 0] = "Hyperlink";
    UrlFieldFormatType[UrlFieldFormatType["Image"] = 1] = "Image";
})(UrlFieldFormatType || (UrlFieldFormatType = {}));
var FieldUserSelectionMode;
(function (FieldUserSelectionMode) {
    FieldUserSelectionMode[FieldUserSelectionMode["PeopleAndGroups"] = 1] = "PeopleAndGroups";
    FieldUserSelectionMode[FieldUserSelectionMode["PeopleOnly"] = 0] = "PeopleOnly";
})(FieldUserSelectionMode || (FieldUserSelectionMode = {}));
var ChoiceFieldFormatType;
(function (ChoiceFieldFormatType) {
    ChoiceFieldFormatType[ChoiceFieldFormatType["Dropdown"] = 0] = "Dropdown";
    ChoiceFieldFormatType[ChoiceFieldFormatType["RadioButtons"] = 1] = "RadioButtons";
})(ChoiceFieldFormatType || (ChoiceFieldFormatType = {}));


/***/ }),

/***/ 7580:
/*!********************************************!*\
  !*** ./node_modules/@pnp/sp/fields/web.js ***!
  \********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 5472);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web, "fields", _types_js__WEBPACK_IMPORTED_MODULE_2__.Fields);
(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web, "availablefields", _types_js__WEBPACK_IMPORTED_MODULE_2__.Fields);


/***/ }),

/***/ 1872:
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/files/folder.js ***!
  \**********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _folders_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../folders/types.js */ 187);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 242);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_folders_types_js__WEBPACK_IMPORTED_MODULE_1__._Folder, "files", _types_js__WEBPACK_IMPORTED_MODULE_2__.Files);


/***/ }),

/***/ 7901:
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/files/index.js ***!
  \*********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _folder_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./folder.js */ 1872);
/* harmony import */ var _item_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./item.js */ 6401);
/* harmony import */ var _web_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./web.js */ 6617);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./types.js */ 242);






/***/ }),

/***/ 6401:
/*!********************************************!*\
  !*** ./node_modules/@pnp/sp/files/item.js ***!
  \********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _items_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../items/types.js */ 132);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 242);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_items_types_js__WEBPACK_IMPORTED_MODULE_1__._Item, "file", _types_js__WEBPACK_IMPORTED_MODULE_2__.File, "file");


/***/ }),

/***/ 3645:
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/sp/files/readable-file.js ***!
  \*****************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   ReadableFile: () => (/* binding */ ReadableFile)
/* harmony export */ });
/* unused harmony export StreamParse */
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../spqueryable.js */ 2678);


function StreamParse() {
    return (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.parseBinderWithErrorCheck)(async (r) => { var _a; return ({ body: r.body, knownLength: parseInt(((_a = r === null || r === void 0 ? void 0 : r.headers) === null || _a === void 0 ? void 0 : _a.get("content-length")) || "-1", 10) }); });
}
class ReadableFile extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPInstance {
    /**
     * Gets the contents of the file as text. Not supported in batching.
     *
     */
    getText() {
        return this.getParsed((0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.TextParse)());
    }
    /**
     * Gets the contents of the file as a blob, does not work in Node.js. Not supported in batching.
     *
     */
    getBlob() {
        return this.getParsed((0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.BlobParse)());
    }
    /**
     * Gets the contents of a file as an ArrayBuffer, works in Node.js. Not supported in batching.
     */
    getBuffer() {
        return this.getParsed((0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.BufferParse)());
    }
    /**
     * Gets the contents of a file as an ArrayBuffer, works in Node.js. Not supported in batching.
     */
    getJSON() {
        return this.getParsed((0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.JSONParse)());
    }
    /**
     * Gets the content of a file as a ReadableStream
     *
     */
    getStream() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.SPQueryable)(this, "$value").using(StreamParse(), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.CacheNever)())((0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.headers)({ "binaryStringResponseBody": "true" }));
    }
    getParsed(parser) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.SPQueryable)(this, "$value").using(parser, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.CacheNever)())();
    }
}


/***/ }),

/***/ 242:
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/files/types.js ***!
  \*********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   File: () => (/* binding */ File),
/* harmony export */   Files: () => (/* binding */ Files),
/* harmony export */   _File: () => (/* binding */ _File),
/* harmony export */   fileFromServerRelativePath: () => (/* binding */ fileFromServerRelativePath)
/* harmony export */ });
/* unused harmony exports _Files, fileFromAbsolutePath, fileFromPath, _Versions, Versions, _Version, Version, CheckinType, MoveOperations, TemplateFileType */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _items_index_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../items/index.js */ 9721);
/* harmony import */ var _utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../utils/odata-url-from.js */ 905);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! ../decorators.js */ 6540);
/* harmony import */ var _utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../utils/extract-web-url.js */ 2997);
/* harmony import */ var _utils_to_resource_path_js__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(/*! ../utils/to-resource-path.js */ 4259);
/* harmony import */ var _utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../utils/encode-path-str.js */ 4729);
/* harmony import */ var _readable_file_js__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ./readable-file.js */ 3645);
/* harmony import */ var _batching_js__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../batching.js */ 8018);
/* harmony import */ var _context_info_index_js__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ../context-info/index.js */ 1734);













/**
 * Describes a collection of File objects
 *
 */
let _Files = class _Files extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_2__._SPCollection {
    /**
     * Gets a File by filename
     *
     * @param name The name of the file, including extension.
     */
    getByUrl(name) {
        if (/%#/.test(name)) {
            throw Error("For file names containing % or # please use web.getFileByServerRelativePath");
        }
        return File(this).concat(`('${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__.encodePath)(name)}')`);
    }
    /**
     * Adds a file using the pound percent safe methods
     *
     * @param url Encoded url of the file
     * @param content The file content
     * @param parameters Additional parameters to control method behavior
     */
    async addUsingPath(url, content, parameters = { Overwrite: false }) {
        const path = [`AddUsingPath(decodedurl='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__.encodePath)(url)}'`];
        if (parameters) {
            if (parameters.Overwrite) {
                path.push(",Overwrite=true");
            }
            if (parameters.EnsureUniqueFileName) {
                path.push(`,EnsureUniqueFileName=${parameters.EnsureUniqueFileName}`);
            }
            if (parameters.AutoCheckoutOnInvalidData) {
                path.push(",AutoCheckoutOnInvalidData=true");
            }
            if (!(0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.stringIsNullOrEmpty)(parameters.XorHash)) {
                path.push(`,XorHash=${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__.encodePath)(parameters.XorHash)}`);
            }
        }
        path.push(")");
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(Files(this, path.join("")), { body: content });
    }
    /**
     * Uploads a file. Not supported for batching
     *
     * @param url The folder-relative url of the file.
     * @param content The Blob file content to add
     * @param props Set of optional values that control the behavior of the underlying addUsingPath and chunkedUpload feature
     * @returns The new File and the raw response.
     */
    async addChunked(url, content, props) {
        // add an empty stub
        const response = await this.addUsingPath(url, null, props);
        const file = fileFromServerRelativePath(this, response.ServerRelativeUrl);
        file.using((0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.CancelAction)(() => {
            return File(file).delete();
        }));
        return file.setContentChunked(content, props);
    }
    /**
     * Adds a ghosted file to an existing list or document library. Not supported for batching.
     *
     * @param fileUrl The server-relative url where you want to save the file.
     * @param templateFileType The type of use to create the file.
     * @returns The template file that was added and the raw response.
     */
    async addTemplateFile(fileUrl, templateFileType) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(Files(this, `addTemplateFile(urloffile='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__.encodePath)(fileUrl)}',templatefiletype=${templateFileType})`));
    }
};
(0,tslib__WEBPACK_IMPORTED_MODULE_10__.__decorate)([
    _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.cancelableScope
], _Files.prototype, "addUsingPath", null);
(0,tslib__WEBPACK_IMPORTED_MODULE_10__.__decorate)([
    _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.cancelableScope
], _Files.prototype, "addChunked", null);
(0,tslib__WEBPACK_IMPORTED_MODULE_10__.__decorate)([
    _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.cancelableScope
], _Files.prototype, "addTemplateFile", null);
_Files = (0,tslib__WEBPACK_IMPORTED_MODULE_10__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_11__.defaultPath)("files")
], _Files);

const Files = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spInvokableFactory)(_Files);
/**
 * Describes a single File instance
 *
 */
class _File extends _readable_file_js__WEBPACK_IMPORTED_MODULE_7__.ReadableFile {
    constructor() {
        super(...arguments);
        this.delete = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.deleteableWithETag)();
    }
    /**
     * Gets a value that specifies the list item field values for the list item corresponding to the file.
     *
     */
    get listItemAllFields() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPInstance)(this, "listItemAllFields");
    }
    /**
     * Gets a collection of versions
     *
     */
    get versions() {
        return Versions(this);
    }
    /**
     * Gets the current locked by user
     *
     */
    async getLockedByUser() {
        const u = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spGet)(File(this, "lockedByUser"));
        if (u["odata.null"] === true) {
            return null;
        }
        else {
            return u;
        }
    }
    /**
     * Approves the file submitted for content approval with the specified comment.
     * Only documents in lists that are enabled for content approval can be approved.
     *
     * @param comment The comment for the approval.
     */
    approve(comment = "") {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(File(this, `approve(comment='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__.encodePath)(comment)}')`));
    }
    /**
     * Stops the chunk upload session without saving the uploaded data. Does not support batching.
     * If the file doesn’t already exist in the library, the partially uploaded file will be deleted.
     * Use this in response to user action (as in a request to cancel an upload) or an error or exception.
     * Use the uploadId value that was passed to the StartUpload method that started the upload session.
     * This method is currently available only on Office 365.
     *
     * @param uploadId The unique identifier of the upload session.
     */
    cancelUpload(uploadId) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(File(this, `cancelUpload(uploadId=guid'${uploadId}')`));
    }
    /**
     * Checks the file in to a document library based on the check-in type.
     *
     * @param comment A comment for the check-in. Its length must be <= 1023.
     * @param checkinType The check-in type for the file.
     */
    checkin(comment = "", checkinType = CheckinType.Major) {
        if (comment.length > 1023) {
            throw Error("The maximum comment length is 1023 characters.");
        }
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(File(this, `checkin(comment='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__.encodePath)(comment)}',checkintype=${checkinType})`));
    }
    /**
     * Checks out the file from a document library.
     */
    checkout() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(File(this, "checkout"));
    }
    /**
     * Copies the file to the destination url.
     *
     * @param url The absolute url or server relative url of the destination file path to copy to.
     * @param shouldOverWrite Should a file with the same name in the same location be overwritten?
     */
    copyTo(url, shouldOverWrite = true) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(File(this, `copyTo(strnewurl='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__.encodePath)(url)}',boverwrite=${shouldOverWrite})`));
    }
    async copyByPath(destUrl, ...rest) {
        let options = {
            ShouldBypassSharedLocks: true,
            ResetAuthorAndCreatedOnCopy: true,
            KeepBoth: false,
        };
        if (rest.length === 2) {
            if (typeof rest[1] === "boolean") {
                options.KeepBoth = rest[1];
            }
            else if (typeof rest[1] === "object") {
                options = { ...options, ...rest[1] };
            }
        }
        return this.moveCopyImpl(destUrl, options, rest[0], "CopyFileByPath");
    }
    /**
     * Denies approval for a file that was submitted for content approval.
     * Only documents in lists that are enabled for content approval can be denied.
     *
     * @param comment The comment for the denial.
     */
    deny(comment = "") {
        if (comment.length > 1023) {
            throw Error("The maximum comment length is 1023 characters.");
        }
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(File(this, `deny(comment='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__.encodePath)(comment)}')`));
    }
    async moveByPath(destUrl, ...rest) {
        let options = {
            KeepBoth: false,
            ShouldBypassSharedLocks: true,
            RetainEditorAndModifiedOnMove: false,
        };
        if (rest.length === 2) {
            if (typeof rest[1] === "boolean") {
                options.KeepBoth = rest[1];
            }
            else if (typeof rest[1] === "object") {
                options = { ...options, ...rest[1] };
            }
        }
        return this.moveCopyImpl(destUrl, options, rest[0], "MoveFileByPath");
    }
    /**
     * Submits the file for content approval with the specified comment.
     *
     * @param comment The comment for the published file. Its length must be <= 1023.
     */
    publish(comment = "") {
        if (comment.length > 1023) {
            throw Error("The maximum comment length is 1023 characters.");
        }
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(File(this, `publish(comment='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__.encodePath)(comment)}')`));
    }
    /**
     * Moves the file to the Recycle Bin and returns the identifier of the new Recycle Bin item.
     *
     * @returns The GUID of the recycled file.
     */
    recycle() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(File(this, "recycle"));
    }
    /**
     * Deletes the file object with options.
     *
     * @param parameters Specifies the options to use when deleting a file.
     */
    async deleteWithParams(parameters) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(File(this, "DeleteWithParameters"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({ parameters }));
    }
    /**
     * Reverts an existing checkout for the file.
     *
     */
    undoCheckout() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(File(this, "undoCheckout"));
    }
    /**
     * Removes the file from content approval or unpublish a major version.
     *
     * @param comment The comment for the unpublish operation. Its length must be <= 1023.
     */
    unpublish(comment = "") {
        if (comment.length > 1023) {
            throw Error("The maximum comment length is 1023 characters.");
        }
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(File(this, `unpublish(comment='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__.encodePath)(comment)}')`));
    }
    /**
     * Checks to see if the file represented by this object exists
     *
     */
    async exists() {
        try {
            const r = await File(this).select("Exists")();
            return r.Exists;
        }
        catch (e) {
            // this treats any error here as the file not existing, which
            // might not be true, but is good enough.
            return false;
        }
    }
    /**
     * Sets the content of a file, for large files use setContentChunked. Not supported in batching.
     *
     * @param content The file content
     *
     */
    async setContent(content) {
        await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(File(this, "$value"), {
            body: content,
            headers: {
                "X-HTTP-Method": "PUT",
            },
        });
        return File(this);
    }
    /**
     * Gets the associated list item for this folder, loading the default properties
     */
    async getItem(...selects) {
        const q = this.listItemAllFields;
        const d = await q.select(...selects)();
        return Object.assign((0,_items_index_js__WEBPACK_IMPORTED_MODULE_3__.Item)([this, (0,_utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_4__.odataUrlFrom)(d)]), d);
    }
    /**
     * Sets the contents of a file using a chunked upload approach. Not supported in batching.
     *
     * @param file The file to upload
     * @param progress A callback function which can be used to track the progress of the upload
     * @param chunkSize The size of each file slice, in bytes (default: 10485760)
     */
    async setContentChunked(file, props) {
        const { progress } = applyChunckedOperationDefaults(props);
        const uploadId = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.getGUID)();
        let first = true;
        let chunk;
        let offset = 0;
        const fileRef = File(this).using((0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.CancelAction)(() => {
            return File(fileRef).cancelUpload(uploadId);
        }));
        const contentStream = sourceToReadableStream(file);
        const reader = contentStream.getReader();
        while ((chunk = await reader.read())) {
            if (chunk.done) {
                progress({ offset, stage: "finishing", uploadId });
                return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(File(fileRef, `finishUpload(uploadId=guid'${uploadId}',fileOffset=${offset})`), { body: (chunk === null || chunk === void 0 ? void 0 : chunk.value) || "" });
            }
            else if (first) {
                progress({ offset, stage: "starting", uploadId });
                offset = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(File(fileRef, `startUpload(uploadId=guid'${uploadId}')`), { body: chunk.value });
                first = false;
            }
            else {
                progress({ offset, stage: "continue", uploadId });
                offset = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(File(fileRef, `continueUpload(uploadId=guid'${uploadId}',fileOffset=${offset})`), { body: chunk.value });
            }
        }
    }
    moveCopyImpl(destUrl, options, overwrite, methodName) {
        // create a timeline we will manipulate for this request
        const poster = File(this);
        // add our pre-request actions, this fixes issues with batching hanging #2668
        poster.on.pre((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.noInherit)(async (url, init, result) => {
            const { ServerRelativeUrl: srcUrl, ["odata.id"]: absoluteUrl } = await File(this).using((0,_batching_js__WEBPACK_IMPORTED_MODULE_8__.BatchNever)()).select("ServerRelativeUrl")();
            const webBaseUrl = new URL((0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_5__.extractWebUrl)(absoluteUrl));
            url = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.combine)(webBaseUrl.toString(), `/_api/SP.MoveCopyUtil.${methodName}(overwrite=@a1)?@a1=${overwrite}`);
            init = (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({
                destPath: (0,_utils_to_resource_path_js__WEBPACK_IMPORTED_MODULE_12__.toResourcePath)((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.isUrlAbsolute)(destUrl) ? destUrl : `${webBaseUrl.protocol}//${webBaseUrl.host}${destUrl}`),
                options,
                srcPath: (0,_utils_to_resource_path_js__WEBPACK_IMPORTED_MODULE_12__.toResourcePath)((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.isUrlAbsolute)(srcUrl) ? srcUrl : `${webBaseUrl.protocol}//${webBaseUrl.host}${srcUrl}`),
            }, init);
            return [url, init, result];
        }));
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(poster).then(() => fileFromPath(this, destUrl));
    }
}
(0,tslib__WEBPACK_IMPORTED_MODULE_10__.__decorate)([
    _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.cancelableScope
], _File.prototype, "copyByPath", null);
(0,tslib__WEBPACK_IMPORTED_MODULE_10__.__decorate)([
    _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.cancelableScope
], _File.prototype, "moveByPath", null);
(0,tslib__WEBPACK_IMPORTED_MODULE_10__.__decorate)([
    _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.cancelableScope
], _File.prototype, "setContentChunked", null);
const File = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spInvokableFactory)(_File);
/**
 * Creates an IFile instance given a base object and a server relative path
 *
 * @param base Valid SPQueryable from which the observers will be used and the web url extracted
 * @param serverRelativePath The server relative url to the file (ex: '/sites/dev/documents/file.txt')
 * @returns IFile instance referencing the file described by the supplied parameters
 */
function fileFromServerRelativePath(base, serverRelativePath) {
    return File([base, (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_5__.extractWebUrl)(base.toUrl())], `_api/web/getFileByServerRelativePath(decodedUrl='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__.encodePath)(serverRelativePath)}')`);
}
/**
 * Creates an IFile instance given a base object and an absolute path
 *
 * @param base Valid SPQueryable from which the observers will be used
 * @param serverRelativePath The absolute url to the file (ex: 'https://tenant.sharepoint.com/sites/dev/documents/file.txt')
 * @returns IFile instance referencing the file described by the supplied parameters
 */
async function fileFromAbsolutePath(base, absoluteFilePath) {
    const { WebFullUrl } = await File(base).using((0,_batching_js__WEBPACK_IMPORTED_MODULE_8__.BatchNever)()).getContextInfo(absoluteFilePath);
    const { pathname } = new URL(absoluteFilePath);
    return fileFromServerRelativePath(File([base, (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.combine)(WebFullUrl, "_api/web")]), decodeURIComponent(pathname));
}
/**
 * Creates an IFile intance given a base object and either an absolute or server relative path to a file
 *
 * @param base Valid SPQueryable from which the observers will be used
 * @param serverRelativePath server relative or absolute url to the file (ex: 'https://tenant.sharepoint.com/sites/dev/documents/file.txt' or '/sites/dev/documents/file.txt')
 * @returns IFile instance referencing the file described by the supplied parameters
 */
async function fileFromPath(base, path) {
    return ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.isUrlAbsolute)(path) ? fileFromAbsolutePath : fileFromServerRelativePath)(base, path);
}
/**
 * Describes a collection of Version objects
 *
 */
let _Versions = class _Versions extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_2__._SPCollection {
    /**
     * Gets a version by id
     *
     * @param versionId The id of the version to retrieve
     */
    getById(versionId) {
        return Version(this).concat(`(${versionId})`);
    }
    /**
     * Deletes all the file version objects in the collection.
     *
     */
    deleteAll() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(Versions(this, "deleteAll"));
    }
    /**
     * Deletes the specified version of the file.
     *
     * @param versionId The ID of the file version to delete.
     */
    deleteById(versionId) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(Versions(this, `deleteById(vid=${versionId})`));
    }
    /**
     * Recycles the specified version of the file.
     *
     * @param versionId The ID of the file version to delete.
     */
    recycleByID(versionId) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(Versions(this, `recycleByID(vid=${versionId})`));
    }
    /**
     * Deletes the file version object with the specified version label.
     *
     * @param label The version label of the file version to delete, for example: 1.2
     */
    deleteByLabel(label) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(Versions(this, `deleteByLabel(versionlabel='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__.encodePath)(label)}')`));
    }
    /**
     * Recycles the file version object with the specified version label.
     *
     * @param label The version label of the file version to delete, for example: 1.2
     */
    recycleByLabel(label) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(Versions(this, `recycleByLabel(versionlabel='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__.encodePath)(label)}')`));
    }
    /**
     * Creates a new file version from the file specified by the version label.
     *
     * @param label The version label of the file version to restore, for example: 1.2
     */
    restoreByLabel(label) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(Versions(this, `restoreByLabel(versionlabel='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__.encodePath)(label)}')`));
    }
};
_Versions = (0,tslib__WEBPACK_IMPORTED_MODULE_10__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_11__.defaultPath)("versions")
], _Versions);

const Versions = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spInvokableFactory)(_Versions);
/**
 * Describes a single Version instance
 *
 */
class _Version extends _readable_file_js__WEBPACK_IMPORTED_MODULE_7__.ReadableFile {
    constructor() {
        super(...arguments);
        this.delete = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.deleteable)();
    }
}
const Version = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spInvokableFactory)(_Version);
/**
 * Types for document check in.
 * Minor = 0
 * Major = 1
 * Overwrite = 2
 */
var CheckinType;
(function (CheckinType) {
    CheckinType[CheckinType["Minor"] = 0] = "Minor";
    CheckinType[CheckinType["Major"] = 1] = "Major";
    CheckinType[CheckinType["Overwrite"] = 2] = "Overwrite";
})(CheckinType || (CheckinType = {}));
/**
 * File move opertions
 */
var MoveOperations;
(function (MoveOperations) {
    /**
     * Produce an error if a file with the same name exists in the destination
     */
    MoveOperations[MoveOperations["None"] = 0] = "None";
    /**
     * Overwrite a file with the same name if it exists. Value is 1.
     */
    MoveOperations[MoveOperations["Overwrite"] = 1] = "Overwrite";
    /**
     * Complete the move operation even if supporting files are separated from the file. Value is 8.
     */
    MoveOperations[MoveOperations["AllowBrokenThickets"] = 8] = "AllowBrokenThickets";
    /**
     * Boolean specifying whether to retain the source of the move's editor and modified by datetime.
     */
    MoveOperations[MoveOperations["RetainEditorAndModifiedOnMove"] = 2048] = "RetainEditorAndModifiedOnMove";
})(MoveOperations || (MoveOperations = {}));
var TemplateFileType;
(function (TemplateFileType) {
    TemplateFileType[TemplateFileType["StandardPage"] = 0] = "StandardPage";
    TemplateFileType[TemplateFileType["WikiPage"] = 1] = "WikiPage";
    TemplateFileType[TemplateFileType["FormPage"] = 2] = "FormPage";
    TemplateFileType[TemplateFileType["ClientSidePage"] = 3] = "ClientSidePage";
})(TemplateFileType || (TemplateFileType = {}));
function applyChunckedOperationDefaults(props) {
    return {
        progress: () => null,
        ...props,
    };
}
/**
 * Converts the source into a ReadableStream we can understand
 */
function sourceToReadableStream(source) {
    if (isBlob(source)) {
        return source.stream();
    }
    else if (hasOn(source)) {
        // we probably have a passthrough stream from NodeFetch or some other type that supports "on(data)"
        return new ReadableStream({
            start(controller) {
                source.on("data", (chunk) => {
                    controller.enqueue(chunk);
                });
                source.on("end", () => {
                    controller.close();
                });
            },
        });
    }
    else if (isBuffer(source)) {
        // we think we have a buffer
        return new ReadableStream({
            start(controller) {
                controller.enqueue(source);
                controller.close();
            },
        });
    }
    else if (isTransform(source)) {
        return source.readable;
    }
    else {
        return source;
    }
}
const NAME = Symbol.toStringTag;
function hasOn(object) {
    // eslint-disable-next-line @typescript-eslint/dot-notation
    return typeof object["on"] === "function";
}
// FROM: node-fetch source code
function isBlob(object) {
    return typeof object === "object" &&
        typeof object.arrayBuffer === "function" &&
        typeof object.type === "string" &&
        typeof object.stream === "function" &&
        typeof object.constructor === "function" &&
        (/^(Blob|File)$/.test(object[NAME]) ||
            /^(Blob|File)$/.test(object.constructor.name));
}
function isBuffer(object) {
    return typeof object === "object" && typeof object.length === "number";
}
function isTransform(object) {
    return typeof object === "object" && typeof object.readable === "object";
}


/***/ }),

/***/ 6617:
/*!*******************************************!*\
  !*** ./node_modules/@pnp/sp/files/web.js ***!
  \*******************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../utils/encode-path-str.js */ 4729);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 242);



_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web.prototype.getFileByServerRelativePath = function (fileRelativeUrl) {
    return (0,_types_js__WEBPACK_IMPORTED_MODULE_2__.fileFromServerRelativePath)(this, fileRelativeUrl);
};
_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web.prototype.getFileById = function (uniqueId) {
    return (0,_types_js__WEBPACK_IMPORTED_MODULE_2__.File)(this, `getFileById('${uniqueId}')`);
};
_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web.prototype.getFileByUrl = function (fileUrl) {
    return (0,_types_js__WEBPACK_IMPORTED_MODULE_2__.File)(this, `getFileByUrl('${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_0__.encodePath)("!@p1::" + fileUrl)}')`);
};


/***/ }),

/***/ 6267:
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/folders/index.js ***!
  \***********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _item_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./item.js */ 1042);
/* harmony import */ var _list_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./list.js */ 6488);
/* harmony import */ var _web_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./web.js */ 6536);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./types.js */ 187);






/***/ }),

/***/ 1042:
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/folders/item.js ***!
  \**********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _items_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../items/types.js */ 132);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 187);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_items_types_js__WEBPACK_IMPORTED_MODULE_1__._Item, "folder", _types_js__WEBPACK_IMPORTED_MODULE_2__.Folder);


/***/ }),

/***/ 6488:
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/folders/list.js ***!
  \**********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _lists_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types.js */ 9706);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 187);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List, "rootFolder", _types_js__WEBPACK_IMPORTED_MODULE_2__.Folder);


/***/ }),

/***/ 187:
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/folders/types.js ***!
  \***********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Folder: () => (/* binding */ Folder),
/* harmony export */   Folders: () => (/* binding */ Folders),
/* harmony export */   _Folder: () => (/* binding */ _Folder),
/* harmony export */   folderFromServerRelativePath: () => (/* binding */ folderFromServerRelativePath)
/* harmony export */ });
/* unused harmony exports _Folders, folderFromAbsolutePath, folderFromPath */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../utils/odata-url-from.js */ 905);
/* harmony import */ var _items_types_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../items/types.js */ 132);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ../decorators.js */ 6540);
/* harmony import */ var _utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../utils/extract-web-url.js */ 2997);
/* harmony import */ var _utils_to_resource_path_js__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! ../utils/to-resource-path.js */ 4259);
/* harmony import */ var _utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../utils/encode-path-str.js */ 4729);
/* harmony import */ var _context_info_index_js__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../context-info/index.js */ 1734);
/* harmony import */ var _batching_js__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../batching.js */ 8018);












let _Folders = class _Folders extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_2__._SPCollection {
    /**
     * Gets a folder by it's name
     *
     * @param name Folder's name
     */
    getByUrl(name) {
        return Folder(this).concat(`('${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__.encodePath)(name)}')`);
    }
    /**
     * Adds a new folder by path and should be prefered over add
     *
     * @param serverRelativeUrl The server relative url of the new folder to create
     * @param overwrite True to overwrite an existing folder, default false
     */
    async addUsingPath(serverRelativeUrl, overwrite = false) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(Folders(this, `addUsingPath(DecodedUrl='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__.encodePath)(serverRelativeUrl)}',overwrite=${overwrite})`));
    }
};
_Folders = (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_10__.defaultPath)("folders")
], _Folders);

const Folders = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spInvokableFactory)(_Folders);
class _Folder extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_2__._SPInstance {
    constructor() {
        super(...arguments);
        this.delete = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.deleteableWithETag)();
    }
    /**
     * Gets this folder's sub folders
     *
     */
    get folders() {
        return Folders(this);
    }
    /**
     * Gets this folder's list item field values
     *
     */
    get listItemAllFields() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPInstance)(this, "listItemAllFields");
    }
    /**
     * Gets the parent folder, if available
     *
     */
    get parentFolder() {
        return Folder(this, "parentFolder");
    }
    /**
     * Gets this folder's properties
     *
     */
    get properties() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPInstance)(this, "properties");
    }
    /**
     * Gets this folder's storage metrics information
     *
     */
    get storageMetrics() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPInstance)(this, "storagemetrics");
    }
    /**
     * Updates folder's properties
     * @param props Folder's properties to update
     */
    async update(props) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPostMerge)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)(props));
    }
    /**
     * Moves the folder to the Recycle Bin and returns the identifier of the new Recycle Bin item.
     */
    recycle() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(Folder(this, "recycle"));
    }
    /**
     * Gets the associated list item for this folder, loading the default properties
     */
    async getItem(...selects) {
        const q = this.listItemAllFields;
        const d = await q.select(...selects)();
        if (d["odata.null"]) {
            throw Error("No associated item was found for this folder. It may be the root folder, which does not have an item.");
        }
        return Object.assign((0,_items_types_js__WEBPACK_IMPORTED_MODULE_4__.Item)([this, (0,_utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_3__.odataUrlFrom)(d)]), d);
    }
    async moveByPath(destUrl, ...rest) {
        let options = {
            KeepBoth: false,
            ShouldBypassSharedLocks: true,
            RetainEditorAndModifiedOnMove: false,
        };
        if (rest.length === 1) {
            if (typeof rest[0] === "boolean") {
                options.KeepBoth = rest[0];
            }
            else if (typeof rest[0] === "object") {
                options = { ...options, ...rest[0] };
            }
        }
        return this.moveCopyImpl(destUrl, options, "MoveFolderByPath");
    }
    async copyByPath(destUrl, ...rest) {
        let options = {
            ShouldBypassSharedLocks: true,
            ResetAuthorAndCreatedOnCopy: true,
            KeepBoth: false,
        };
        if (rest.length === 1) {
            if (typeof rest[0] === "boolean") {
                options.KeepBoth = rest[0];
            }
            else if (typeof rest[0] === "object") {
                options = { ...options, ...rest[0] };
            }
        }
        return this.moveCopyImpl(destUrl, options, "CopyFolderByPath");
    }
    /**
     * Deletes the folder object with options.
     *
     * @param parameters Specifies the options to use when deleting a folder.
     */
    async deleteWithParams(parameters) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(Folder(this, "DeleteWithParameters"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)({ parameters }));
    }
    /**
     * Create the subfolder inside the current folder, as specified by the leafPath
     *
     * @param leafPath leafName of the new folder
     */
    async addSubFolderUsingPath(leafPath) {
        await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(Folder(this, "AddSubFolderUsingPath"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)({ leafPath: (0,_utils_to_resource_path_js__WEBPACK_IMPORTED_MODULE_11__.toResourcePath)(leafPath) }));
        return this.folders.getByUrl(leafPath);
    }
    /**
     * Gets the parent information for this folder's list and web
     */
    async getParentInfos() {
        const urlInfo = await this.select("ServerRelativeUrl", "ListItemAllFields/ParentList/Id", "ListItemAllFields/ParentList/RootFolder/UniqueId", "ListItemAllFields/ParentList/RootFolder/ServerRelativeUrl", "ListItemAllFields/ParentList/RootFolder/ServerRelativePath", "ListItemAllFields/ParentList/ParentWeb/Id", "ListItemAllFields/ParentList/ParentWeb/Url", "ListItemAllFields/ParentList/ParentWeb/ServerRelativeUrl", "ListItemAllFields/ParentList/ParentWeb/ServerRelativePath").expand("ListItemAllFields/ParentList", "ListItemAllFields/ParentList/RootFolder", "ListItemAllFields/ParentList/ParentWeb")();
        return {
            Folder: {
                ServerRelativeUrl: urlInfo.ServerRelativeUrl,
            },
            ParentList: {
                Id: urlInfo.ListItemAllFields.ParentList.Id,
                RootFolderServerRelativePath: urlInfo.ListItemAllFields.ParentList.RootFolder.ServerRelativePath,
                RootFolderServerRelativeUrl: urlInfo.ListItemAllFields.ParentList.RootFolder.ServerRelativeUrl,
                RootFolderUniqueId: urlInfo.ListItemAllFields.ParentList.RootFolder.UniqueId,
            },
            ParentWeb: {
                Id: urlInfo.ListItemAllFields.ParentList.ParentWeb.Id,
                ServerRelativePath: urlInfo.ListItemAllFields.ParentList.ParentWeb.ServerRelativePath,
                ServerRelativeUrl: urlInfo.ListItemAllFields.ParentList.ParentWeb.ServerRelativeUrl,
                Url: urlInfo.ListItemAllFields.ParentList.ParentWeb.Url,
            },
        };
    }
    /**
     * Implementation of folder move/copy
     *
     * @param destUrl The server relative path to which the folder will be copied/moved
     * @param options Any options
     * @param methodName The method to call
     * @returns An IFolder representing the moved or copied folder
     */
    moveCopyImpl(destUrl, options, methodName) {
        // create a timeline we will manipulate for this request
        const poster = Folder(this);
        // add our pre-request actions, this fixes issues with batching hanging #2668
        poster.on.pre((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.noInherit)(async (url, init, result) => {
            const { ServerRelativeUrl: srcUrl, ["odata.id"]: absoluteUrl } = await Folder(this).using((0,_batching_js__WEBPACK_IMPORTED_MODULE_8__.BatchNever)()).select("ServerRelativeUrl")();
            const uri = new URL((0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_5__.extractWebUrl)(absoluteUrl));
            url = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.combine)(uri.href, `/_api/SP.MoveCopyUtil.${methodName}()`);
            init = (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)({
                destPath: (0,_utils_to_resource_path_js__WEBPACK_IMPORTED_MODULE_11__.toResourcePath)((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isUrlAbsolute)(destUrl) ? destUrl : (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.combine)(uri.origin, destUrl)),
                options,
                srcPath: (0,_utils_to_resource_path_js__WEBPACK_IMPORTED_MODULE_11__.toResourcePath)((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.combine)(uri.origin, srcUrl)),
            }, init);
            return [url, init, result];
        }));
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(poster).then(() => folderFromPath(this, destUrl));
    }
}
(0,tslib__WEBPACK_IMPORTED_MODULE_9__.__decorate)([
    _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.cancelableScope
], _Folder.prototype, "moveByPath", null);
(0,tslib__WEBPACK_IMPORTED_MODULE_9__.__decorate)([
    _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.cancelableScope
], _Folder.prototype, "copyByPath", null);
const Folder = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spInvokableFactory)(_Folder);
/**
 * Creates an IFolder instance given a base object and a server relative path
 *
 * @param base Valid SPQueryable from which the observers will be used and the web url extracted
 * @param serverRelativePath The server relative url to the folder (ex: '/sites/dev/documents/folder3')
 * @returns IFolder instance referencing the folder described by the supplied parameters
 */
function folderFromServerRelativePath(base, serverRelativePath) {
    return Folder([base, (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_5__.extractWebUrl)(base.toUrl())], `_api/web/getFolderByServerRelativePath(decodedUrl='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_6__.encodePath)(serverRelativePath)}')`);
}
/**
 * Creates an IFolder instance given a base object and an absolute path
 *
 * @param base Valid SPQueryable from which the observers will be used
 * @param serverRelativePath The absolute url to the folder (ex: 'https://tenant.sharepoint.com/sites/dev/documents/folder/')
 * @returns IFolder instance referencing the folder described by the supplied parameters
 */
async function folderFromAbsolutePath(base, absoluteFolderPath) {
    const { WebFullUrl } = await Folder(base).using((0,_batching_js__WEBPACK_IMPORTED_MODULE_8__.BatchNever)()).getContextInfo(absoluteFolderPath);
    const { pathname } = new URL(absoluteFolderPath);
    return folderFromServerRelativePath(Folder([base, (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.combine)(WebFullUrl, "_api/web")]), decodeURIComponent(pathname));
}
/**
 * Creates an IFolder intance given a base object and either an absolute or server relative path to a folder
 *
 * @param base Valid SPQueryable from which the observers will be used
 * @param serverRelativePath server relative or absolute url to the file (ex: 'https://tenant.sharepoint.com/sites/dev/documents/folder' or '/sites/dev/documents/folder')
 * @returns IFile instance referencing the file described by the supplied parameters
 */
async function folderFromPath(base, path) {
    return ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isUrlAbsolute)(path) ? folderFromAbsolutePath : folderFromServerRelativePath)(base, path);
}


/***/ }),

/***/ 6536:
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/folders/web.js ***!
  \*********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 187);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web, "folders", _types_js__WEBPACK_IMPORTED_MODULE_2__.Folders);
(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web, "rootFolder", _types_js__WEBPACK_IMPORTED_MODULE_2__.Folder);
_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web.prototype.getFolderByServerRelativePath = function (folderRelativeUrl) {
    return (0,_types_js__WEBPACK_IMPORTED_MODULE_2__.folderFromServerRelativePath)(this, folderRelativeUrl);
};
_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web.prototype.getFolderById = function (uniqueId) {
    return (0,_types_js__WEBPACK_IMPORTED_MODULE_2__.Folder)(this, `getFolderById('${uniqueId}')`);
};


/***/ }),

/***/ 8744:
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/forms/index.js ***!
  \*********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _list_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./list.js */ 4777);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 282);




/***/ }),

/***/ 4777:
/*!********************************************!*\
  !*** ./node_modules/@pnp/sp/forms/list.js ***!
  \********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _lists_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types.js */ 9706);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 282);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List, "forms", _types_js__WEBPACK_IMPORTED_MODULE_2__.Forms);


/***/ }),

/***/ 282:
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/forms/types.js ***!
  \*********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Forms: () => (/* binding */ Forms)
/* harmony export */ });
/* unused harmony exports _Forms, _Form, Form */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../decorators.js */ 6540);



/**
 * Describes a collection of Form objects
 *
 */
let _Forms = class _Forms extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPCollection {
    /**
     * Gets a form by id
     *
     * @param id The guid id of the item to retrieve
     */
    getById(id) {
        return Form(this).concat(`('${id}')`);
    }
};
_Forms = (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_2__.defaultPath)("forms")
], _Forms);

const Forms = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_Forms);
/**
 * Describes a single of Form instance
 *
 */
class _Form extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPInstance {
}
const Form = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_Form);


/***/ }),

/***/ 5998:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/hubsites/index.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _fi_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../fi.js */ 6553);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 3175);
/* harmony import */ var _site_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./site.js */ 7316);
/* harmony import */ var _web_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./web.js */ 6821);





Reflect.defineProperty(_fi_js__WEBPACK_IMPORTED_MODULE_0__.SPFI.prototype, "hubSites", {
    configurable: true,
    enumerable: true,
    get: function () {
        return this.create(_types_js__WEBPACK_IMPORTED_MODULE_1__.HubSites);
    },
});


/***/ }),

/***/ 7316:
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/hubsites/site.js ***!
  \***********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _sites_types_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../sites/types.js */ 1114);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../spqueryable.js */ 2678);


_sites_types_js__WEBPACK_IMPORTED_MODULE_0__._Site.prototype.joinHubSite = async function (siteId) {
    await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)((0,_sites_types_js__WEBPACK_IMPORTED_MODULE_0__.Site)(this, `joinHubSite('${siteId}')`));
};
_sites_types_js__WEBPACK_IMPORTED_MODULE_0__._Site.prototype.registerHubSite = async function () {
    await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)((0,_sites_types_js__WEBPACK_IMPORTED_MODULE_0__.Site)(this, "registerHubSite"));
};
_sites_types_js__WEBPACK_IMPORTED_MODULE_0__._Site.prototype.unRegisterHubSite = async function () {
    await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)((0,_sites_types_js__WEBPACK_IMPORTED_MODULE_0__.Site)(this, "unRegisterHubSite"));
};


/***/ }),

/***/ 3175:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/hubsites/types.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   HubSites: () => (/* binding */ HubSites)
/* harmony export */ });
/* unused harmony exports _HubSites, _HubSite, HubSite */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../decorators.js */ 6540);
/* harmony import */ var _sites_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sites/types.js */ 1114);




let _HubSites = class _HubSites extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPCollection {
    /**
     * Gets a Hub Site from the collection by id
     *
     * @param id The Id of the Hub Site
     */
    getById(id) {
        return HubSite(this, `GetById?hubSiteId='${id}'`);
    }
};
_HubSites = (0,tslib__WEBPACK_IMPORTED_MODULE_2__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_3__.defaultPath)("_api/hubsites")
], _HubSites);

const HubSites = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_HubSites);
class _HubSite extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPInstance {
    /**
     * Gets the ISite instance associated with this hubsite
     */
    async getSite() {
        const d = await this.select("SiteUrl")();
        return (0,_sites_types_js__WEBPACK_IMPORTED_MODULE_1__.Site)([this, d.SiteUrl]);
    }
}
const HubSite = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_HubSite);


/***/ }),

/***/ 6821:
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/hubsites/web.js ***!
  \**********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../spqueryable.js */ 2678);


_webs_types_js__WEBPACK_IMPORTED_MODULE_0__._Web.prototype.hubSiteData = async function (forceRefresh = false) {
    const data = await (0,_webs_types_js__WEBPACK_IMPORTED_MODULE_0__.Web)(this, `hubSiteData(${forceRefresh})`)();
    if (typeof data === "string") {
        return JSON.parse(data);
    }
    return data;
};
_webs_types_js__WEBPACK_IMPORTED_MODULE_0__._Web.prototype.syncHubSiteTheme = function () {
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)((0,_webs_types_js__WEBPACK_IMPORTED_MODULE_0__.Web)(this, "syncHubSiteTheme"));
};


/***/ }),

/***/ 7881:
/*!***************************************!*\
  !*** ./node_modules/@pnp/sp/index.js ***!
  \***************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   SPCollection: () => (/* reexport safe */ _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.SPCollection),
/* harmony export */   SPFx: () => (/* reexport safe */ _behaviors_spfx_js__WEBPACK_IMPORTED_MODULE_10__.SPFx),
/* harmony export */   extractWebUrl: () => (/* reexport safe */ _utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_3__.extractWebUrl),
/* harmony export */   spfi: () => (/* reexport safe */ _fi_js__WEBPACK_IMPORTED_MODULE_1__.spfi)
/* harmony export */ });
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./spqueryable.js */ 2678);
/* harmony import */ var _fi_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./fi.js */ 6553);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 8713);
/* harmony import */ var _utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./utils/extract-web-url.js */ 2997);
/* harmony import */ var _utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./utils/odata-url-from.js */ 905);
/* harmony import */ var _utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./utils/encode-path-str.js */ 4729);
/* harmony import */ var _behaviors_defaults_js__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./behaviors/defaults.js */ 7140);
/* harmony import */ var _behaviors_telemetry_js__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ./behaviors/telemetry.js */ 2630);
/* harmony import */ var _behaviors_request_digest_js__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ./behaviors/request-digest.js */ 670);
/* harmony import */ var _behaviors_spbrowser_js__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ./behaviors/spbrowser.js */ 6793);
/* harmony import */ var _behaviors_spfx_js__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ./behaviors/spfx.js */ 4243);

















/***/ }),

/***/ 9721:
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/items/index.js ***!
  \*********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Item: () => (/* reexport safe */ _types_js__WEBPACK_IMPORTED_MODULE_1__.Item)
/* harmony export */ });
/* harmony import */ var _list_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./list.js */ 5685);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 132);




/***/ }),

/***/ 5685:
/*!********************************************!*\
  !*** ./node_modules/@pnp/sp/items/list.js ***!
  \********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _lists_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types.js */ 9706);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 132);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List, "items", _types_js__WEBPACK_IMPORTED_MODULE_2__.Items);


/***/ }),

/***/ 132:
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/items/types.js ***!
  \*********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Item: () => (/* binding */ Item),
/* harmony export */   Items: () => (/* binding */ Items),
/* harmony export */   _Item: () => (/* binding */ _Item)
/* harmony export */ });
/* unused harmony exports _Items, _ItemVersions, ItemVersions, _ItemVersion, ItemVersion */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _pnp_sp__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/sp */ 7881);
/* harmony import */ var _lists_types_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../lists/types.js */ 9706);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../decorators.js */ 6540);







/**
 * Describes a collection of Item objects
 *
 */
let _Items = class _Items extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPCollection {
    /**
    * Gets an Item by id
    *
    * @param id The integer id of the item to retrieve
    */
    getById(id) {
        return Item(this).concat(`(${id})`);
    }
    /**
     * Gets BCS Item by string id
     *
     * @param stringId The string id of the BCS item to retrieve
     */
    getItemByStringId(stringId) {
        // creates an item with the parent list path and append out method call
        return Item([this, this.parentUrl], `getItemByStringId('${stringId}')`);
    }
    /**
     * Skips the specified number of items (https://msdn.microsoft.com/en-us/library/office/fp142385.aspx#sectionSection6)
     *
     * @param skip The starting id where the page should start, use with top to specify pages
     * @param reverse It true the PagedPrev=true parameter is added allowing backwards navigation in the collection
     */
    skip(skip, reverse = false) {
        if (reverse) {
            this.query.set("$skiptoken", `Paged=TRUE&PagedPrev=TRUE&p_ID=${skip}`);
        }
        else {
            this.query.set("$skiptoken", `Paged=TRUE&p_ID=${skip}`);
        }
        return this;
    }
    [Symbol.asyncIterator]() {
        const nextInit = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.SPCollection)(this).using((0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_4__.parseBinderWithErrorCheck)(async (r) => {
            const json = await r.json();
            const nextLink = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(json, "d") && (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(json.d, "__next") ? json.d.__next : json["odata.nextLink"];
            return {
                hasNext: typeof nextLink === "string" && nextLink.length > 0,
                nextLink,
                value: (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_4__.parseODataJSON)(json),
            };
        }));
        const queryParams = ["$top", "$select", "$expand", "$filter", "$orderby", "$skiptoken"];
        for (let i = 0; i < queryParams.length; i++) {
            const param = this.query.get(queryParams[i]);
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.objectDefinedNotNull)(param)) {
                nextInit.query.set(queryParams[i], param);
            }
        }
        return {
            _next: nextInit,
            async next() {
                if (this._next === null) {
                    return { done: true, value: undefined };
                }
                const result = await this._next();
                if (result.hasNext) {
                    this._next = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.SPCollection)([this._next, result.nextLink]);
                    return { done: false, value: result.value };
                }
                else {
                    this._next = null;
                    return { done: false, value: result.value };
                }
            },
        };
    }
    /**
     * Adds a new item to the collection
     *
     * @param properties The new items's properties
     * @param listItemEntityTypeFullName The type name of the list's entities
     */
    async add(properties = {}) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_4__.body)(properties));
    }
};
_Items = (0,tslib__WEBPACK_IMPORTED_MODULE_5__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_6__.defaultPath)("items")
], _Items);

const Items = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_Items);
/**
 * Descrines a single Item instance
 *
 */
class _Item extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPInstance {
    constructor() {
        super(...arguments);
        this.delete = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.deleteableWithETag)();
    }
    /**
     * Gets the effective base permissions for the item
     *
     */
    get effectiveBasePermissions() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.SPQueryable)(this, "EffectiveBasePermissions");
    }
    /**
     * Gets the effective base permissions for the item in a UI context
     *
     */
    get effectiveBasePermissionsForUI() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.SPQueryable)(this, "EffectiveBasePermissionsForUI");
    }
    /**
     * Gets the field values for this list item in their HTML representation
     *
     */
    get fieldValuesAsHTML() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.SPInstance)(this, "FieldValuesAsHTML");
    }
    /**
     * Gets the field values for this list item in their text representation
     *
     */
    get fieldValuesAsText() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.SPInstance)(this, "FieldValuesAsText");
    }
    /**
     * Gets the field values for this list item for use in editing controls
     *
     */
    get fieldValuesForEdit() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.SPInstance)(this, "FieldValuesForEdit");
    }
    /**
     * Gets the collection of versions associated with this item
     */
    get versions() {
        return ItemVersions(this);
    }
    /**
     * this item's list
     */
    get list() {
        return this.getParent(_lists_types_js__WEBPACK_IMPORTED_MODULE_3__.List, "", this.parentUrl.substring(0, this.parentUrl.lastIndexOf("/")));
    }
    /**
     * Updates this list instance with the supplied properties
     *
     * @param properties A plain object hash of values to update for the list
     * @param eTag Value used in the IF-Match header, by default "*"
     */
    async update(properties, eTag = "*") {
        const postBody = (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_4__.body)(properties, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_4__.headers)({
            "IF-Match": eTag,
            "X-HTTP-Method": "MERGE",
        }));
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(Item(this).using(ItemUpdatedParser()), postBody);
    }
    /**
     * Moves the list item to the Recycle Bin and returns the identifier of the new Recycle Bin item.
     */
    recycle() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(Item(this, "recycle"));
    }
    /**
     * Deletes the item object with options.
     *
     * @param parameters Specifies the options to use when deleting a item.
     */
    async deleteWithParams(parameters) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(Item(this, "DeleteWithParameters"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_4__.body)({ parameters }));
    }
    /**
     * Gets a string representation of the full URL to the WOPI frame.
     * If there is no associated WOPI application, or no associated action, an empty string is returned.
     *
     * @param action Display mode: 0: view, 1: edit, 2: mobileView, 3: interactivePreview
     */
    async getWopiFrameUrl(action = 0) {
        const i = Item(this, "getWOPIFrameUrl(@action)");
        i.query.set("@action", action);
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(i);
    }
    /**
     * Validates and sets the values of the specified collection of fields for the list item.
     *
     * @param formValues The fields to change and their new values.
     * @param bNewDocumentUpdate true if the list item is a document being updated after upload; otherwise false.
     */
    validateUpdateListItem(formValues, bNewDocumentUpdate = false) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(Item(this, "validateupdatelistitem"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_4__.body)({ formValues, bNewDocumentUpdate }));
    }
    /**
     * Gets the parent information for this item's list and web
     */
    async getParentInfos() {
        const urlInfo = await this.select("Id", "ParentList/Id", "ParentList/Title", "ParentList/RootFolder/UniqueId", "ParentList/RootFolder/ServerRelativeUrl", "ParentList/RootFolder/ServerRelativePath", "ParentList/ParentWeb/Id", "ParentList/ParentWeb/Url", "ParentList/ParentWeb/ServerRelativeUrl", "ParentList/ParentWeb/ServerRelativePath").expand("ParentList", "ParentList/RootFolder", "ParentList/ParentWeb")();
        return {
            Item: {
                Id: urlInfo.Id,
            },
            ParentList: {
                Id: urlInfo.ParentList.Id,
                Title: urlInfo.ParentList.Title,
                RootFolderServerRelativePath: urlInfo.ParentList.RootFolder.ServerRelativePath,
                RootFolderServerRelativeUrl: urlInfo.ParentList.RootFolder.ServerRelativeUrl,
                RootFolderUniqueId: urlInfo.ParentList.RootFolder.UniqueId,
            },
            ParentWeb: {
                Id: urlInfo.ParentList.ParentWeb.Id,
                ServerRelativePath: urlInfo.ParentList.ParentWeb.ServerRelativePath,
                ServerRelativeUrl: urlInfo.ParentList.ParentWeb.ServerRelativeUrl,
                Url: urlInfo.ParentList.ParentWeb.Url,
            },
        };
    }
    async setImageField(fieldName, imageName, imageContent) {
        const contextInfo = await this.getParentInfos();
        const webUrl = (0,_pnp_sp__WEBPACK_IMPORTED_MODULE_2__.extractWebUrl)(this.toUrl());
        const q = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.SPQueryable)([this, webUrl], "/_api/web/UploadImage");
        q.concat("(listTitle=@a1,imageName=@a2,listId=@a3,itemId=@a4)");
        q.query.set("@a1", `'${contextInfo.ParentList.Title}'`);
        q.query.set("@a2", `'${imageName}'`);
        q.query.set("@a3", `'${contextInfo.ParentList.Id}'`);
        q.query.set("@a4", contextInfo.Item.Id);
        const result = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(q, { body: imageContent });
        const itemInfo = {
            "type": "thumbnail",
            "fileName": result.Name,
            "nativeFile": {},
            "fieldName": fieldName,
            "serverUrl": contextInfo.ParentWeb.Url.replace(contextInfo.ParentWeb.ServerRelativeUrl, ""),
            "serverRelativeUrl": result.ServerRelativeUrl,
            "id": result.UniqueId,
        };
        return this.validateUpdateListItem([{
                FieldName: fieldName,
                FieldValue: JSON.stringify(itemInfo),
            }]);
    }
}
const Item = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_Item);
/**
 * Describes a collection of Version objects
 *
 */
let _ItemVersions = class _ItemVersions extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPCollection {
    /**
     * Gets a version by id
     *
     * @param versionId The id of the version to retrieve
     */
    getById(versionId) {
        return ItemVersion(this).concat(`(${versionId})`);
    }
};
_ItemVersions = (0,tslib__WEBPACK_IMPORTED_MODULE_5__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_6__.defaultPath)("versions")
], _ItemVersions);

const ItemVersions = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_ItemVersions);
/**
 * Describes a single Version instance
 *
 */
class _ItemVersion extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPInstance {
    constructor() {
        super(...arguments);
        this.delete = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.deleteableWithETag)();
    }
}
const ItemVersion = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_ItemVersion);
function ItemUpdatedParser() {
    return (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_4__.parseBinderWithErrorCheck)(async (r) => ({
        etag: r.headers.get("etag"),
    }));
}


/***/ }),

/***/ 2033:
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/lists/index.js ***!
  \*********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _web_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web.js */ 2519);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 9706);




/***/ }),

/***/ 9706:
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/lists/types.js ***!
  \*********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   List: () => (/* binding */ List),
/* harmony export */   Lists: () => (/* binding */ Lists),
/* harmony export */   _List: () => (/* binding */ _List)
/* harmony export */ });
/* unused harmony exports _Lists, RenderListDataOptions, ControlMode */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../utils/odata-url-from.js */ 905);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../decorators.js */ 6540);
/* harmony import */ var _utils_to_resource_path_js__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../utils/to-resource-path.js */ 4259);
/* harmony import */ var _utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../utils/encode-path-str.js */ 4729);








let _Lists = class _Lists extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_2__._SPCollection {
    /**
     * Gets a list from the collection by guid id
     *
     * @param id The Id of the list (GUID)
     */
    getById(id) {
        return List(this).concat(`('${id}')`);
    }
    /**
     * Gets a list from the collection by title
     *
     * @param title The title of the list
     */
    getByTitle(title) {
        return List(this, `getByTitle('${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_4__.encodePath)(title)}')`);
    }
    /**
     * Adds a new list to the collection
     *
     * @param title The new list's title
     * @param description The new list's description
     * @param template The list template value
     * @param enableContentTypes If true content types will be allowed and enabled, otherwise they will be disallowed and not enabled
     * @param additionalSettings Will be passed as part of the list creation body
     */
    async add(title, desc = "", template = 100, enableContentTypes = false, additionalSettings = {}) {
        const addSettings = {
            "AllowContentTypes": enableContentTypes,
            "BaseTemplate": template,
            "ContentTypesEnabled": enableContentTypes,
            "Description": desc,
            "Title": title,
            ...additionalSettings,
        };
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)(addSettings));
    }
    /**
     * Ensures that the specified list exists in the collection (note: this method not supported for batching)
     *
     * @param title The new list's title
     * @param desc The new list's description
     * @param template The list template value
     * @param enableContentTypes If true content types will be allowed and enabled, otherwise they will be disallowed and not enabled
     * @param additionalSettings Will be passed as part of the list creation body or used to update an existing list
     */
    async ensure(title, desc = "", template = 100, enableContentTypes = false, additionalSettings = {}) {
        const addOrUpdateSettings = { Title: title, Description: desc, ContentTypesEnabled: enableContentTypes, ...additionalSettings };
        const list = this.getByTitle(addOrUpdateSettings.Title);
        try {
            await list.select("Title")();
            const data = await list.update(addOrUpdateSettings);
            return { created: false, data, list: this.getByTitle(addOrUpdateSettings.Title) };
        }
        catch (e) {
            const data = await this.add(title, desc, template, enableContentTypes, addOrUpdateSettings);
            return { created: true, data, list: this.getByTitle(addOrUpdateSettings.Title) };
        }
    }
    /**
     * Gets a list that is the default asset location for images or other files, which the users upload to their wiki pages.
     */
    async ensureSiteAssetsLibrary() {
        const json = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(Lists(this, "ensuresiteassetslibrary"));
        return List([this, (0,_utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_3__.odataUrlFrom)(json)]);
    }
    /**
     * Gets a list that is the default location for wiki pages.
     */
    async ensureSitePagesLibrary() {
        const json = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(Lists(this, "ensuresitepageslibrary"));
        return List([this, (0,_utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_3__.odataUrlFrom)(json)]);
    }
};
_Lists = (0,tslib__WEBPACK_IMPORTED_MODULE_5__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_6__.defaultPath)("lists")
], _Lists);

const Lists = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spInvokableFactory)(_Lists);
class _List extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_2__._SPInstance {
    constructor() {
        super(...arguments);
        this.delete = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.deleteableWithETag)();
    }
    /**
     * Gets the effective base permissions of this list
     *
     */
    get effectiveBasePermissions() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPQueryable)(this, "EffectiveBasePermissions");
    }
    /**
     * Gets the event receivers attached to this list
     *
     */
    get eventReceivers() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPCollection)(this, "EventReceivers");
    }
    /**
     * Gets the related fields of this list
     *
     */
    get relatedFields() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPQueryable)(this, "getRelatedFields");
    }
    /**
     * Gets the IRM settings for this list
     *
     */
    get informationRightsManagementSettings() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPQueryable)(this, "InformationRightsManagementSettings");
    }
    /**
     * Updates this list intance with the supplied properties
     *
     * @param properties A plain object hash of values to update for the list
     * @param eTag Value used in the IF-Match header, by default "*"
     */
    async update(properties, eTag = "*") {
        const data = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPostMerge)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)(properties, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.headers)({ "IF-Match": eTag })));
        return data;
    }
    /**
     * Returns the collection of changes from the change log that have occurred within the list, based on the specified query.
     * @param query A query that is performed against the change log.
     */
    getChanges(query) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(List(this, "getchanges"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)({ query }));
    }
    /**
     * Returns the collection of items in the list based on the provided CamlQuery
     * @param query A query that is performed against the list
     * @param expands An expanded array of n items that contains fields to expand in the CamlQuery
     */
    getItemsByCAMLQuery(query, ...expands) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(List(this, "getitems").expand(...expands), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)({ query }));
    }
    /**
     * See: https://msdn.microsoft.com/en-us/library/office/dn292554.aspx
     * @param query An object that defines the change log item query
     */
    getListItemChangesSinceToken(query) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(List(this, "getlistitemchangessincetoken").using((0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.TextParse)()), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)({ query }));
    }
    /**
     * Moves the list to the Recycle Bin and returns the identifier of the new Recycle Bin item.
     */
    async recycle() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(List(this, "recycle"));
    }
    /**
     * Renders list data based on the view xml provided
     * @param viewXml A string object representing a view xml
     */
    async renderListData(viewXml) {
        const q = List(this, "renderlistdata(@viewXml)");
        q.query.set("@viewXml", `'${viewXml}'`);
        const data = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(q);
        return JSON.parse(data);
    }
    /**
     * Returns the data for the specified query view
     *
     * @param parameters The parameters to be used to render list data as JSON string.
     * @param overrideParams The parameters that are used to override and extend the regular SPRenderListDataParameters.
     * @param query Allows setting of query parameters
     */
    // eslint-disable-next-line max-len
    renderListDataAsStream(parameters, overrideParameters = null, query = new Map()) {
        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(parameters, "RenderOptions") && (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isArray)(parameters.RenderOptions)) {
            parameters.RenderOptions = parameters.RenderOptions.reduce((v, c) => v + c);
        }
        const clone = List(this, "RenderListDataAsStream");
        if (query && query.size > 0) {
            query.forEach((v, k) => clone.query.set(k, v));
        }
        const params = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.objectDefinedNotNull)(overrideParameters) ? { parameters, overrideParameters } : { parameters };
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(clone, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)(params));
    }
    /**
     * Gets the field values and field schema attributes for a list item.
     * @param itemId Item id of the item to render form data for
     * @param formId The id of the form
     * @param mode Enum representing the control mode of the form (Display, Edit, New)
     */
    async renderListFormData(itemId, formId, mode) {
        const data = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(List(this, `renderlistformdata(itemid=${itemId}, formid='${formId}', mode='${mode}')`));
        // data will be a string, so we parse it again
        return JSON.parse(data);
    }
    /**
     * Reserves a list item ID for idempotent list item creation.
     */
    async reserveListItemId() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(List(this, "reservelistitemid"));
    }
    /**
     * Creates an item using path (in a folder), validates and sets its field values.
     *
     * @param formValues The fields to change and their new values.
     * @param decodedUrl Path decoded url; folder's server relative path.
     * @param bNewDocumentUpdate true if the list item is a document being updated after upload; otherwise false.
     * @param checkInComment Optional check in comment.
     * @param additionalProps Optional set of additional properties LeafName new document file name,
     */
    async addValidateUpdateItemUsingPath(formValues, decodedUrl, bNewDocumentUpdate = false, checkInComment, additionalProps) {
        const addProps = {
            FolderPath: (0,_utils_to_resource_path_js__WEBPACK_IMPORTED_MODULE_7__.toResourcePath)(decodedUrl),
        };
        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.objectDefinedNotNull)(additionalProps)) {
            if (additionalProps.leafName) {
                addProps.LeafName = (0,_utils_to_resource_path_js__WEBPACK_IMPORTED_MODULE_7__.toResourcePath)(additionalProps.leafName);
            }
            if (additionalProps.objectType) {
                addProps.UnderlyingObjectType = additionalProps.objectType;
            }
        }
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(List(this, "AddValidateUpdateItemUsingPath()"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)({
            bNewDocumentUpdate,
            checkInComment,
            formValues,
            listItemCreateInfo: addProps,
        }));
    }
    /**
     * Gets the parent information for this item's list and web
     */
    async getParentInfos() {
        const urlInfo = await this.select("Id", "RootFolder/UniqueId", "RootFolder/ServerRelativeUrl", "RootFolder/ServerRelativePath", "ParentWeb/Id", "ParentWeb/Url", "ParentWeb/ServerRelativeUrl", "ParentWeb/ServerRelativePath").expand("RootFolder", "ParentWeb")();
        return {
            List: {
                Id: urlInfo.Id,
                RootFolderServerRelativePath: urlInfo.RootFolder.ServerRelativePath,
                RootFolderServerRelativeUrl: urlInfo.RootFolder.ServerRelativeUrl,
                RootFolderUniqueId: urlInfo.RootFolder.UniqueId,
            },
            ParentWeb: {
                Id: urlInfo.ParentWeb.Id,
                ServerRelativePath: urlInfo.ParentWeb.ServerRelativePath,
                ServerRelativeUrl: urlInfo.ParentWeb.ServerRelativeUrl,
                Url: urlInfo.ParentWeb.Url,
            },
        };
    }
}
const List = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spInvokableFactory)(_List);
/**
 * Enum representing the options of the RenderOptions property on IRenderListDataParameters interface
 */
var RenderListDataOptions;
(function (RenderListDataOptions) {
    RenderListDataOptions[RenderListDataOptions["None"] = 0] = "None";
    RenderListDataOptions[RenderListDataOptions["ContextInfo"] = 1] = "ContextInfo";
    RenderListDataOptions[RenderListDataOptions["ListData"] = 2] = "ListData";
    RenderListDataOptions[RenderListDataOptions["ListSchema"] = 4] = "ListSchema";
    RenderListDataOptions[RenderListDataOptions["MenuView"] = 8] = "MenuView";
    RenderListDataOptions[RenderListDataOptions["ListContentType"] = 16] = "ListContentType";
    /**
     * The returned list will have a FileSystemItemId field on each item if possible.
     */
    RenderListDataOptions[RenderListDataOptions["FileSystemItemId"] = 32] = "FileSystemItemId";
    /**
     * Returns the client form schema to add and edit items.
     */
    RenderListDataOptions[RenderListDataOptions["ClientFormSchema"] = 64] = "ClientFormSchema";
    /**
     * Returns QuickLaunch navigation nodes.
     */
    RenderListDataOptions[RenderListDataOptions["QuickLaunch"] = 128] = "QuickLaunch";
    /**
     * Returns Spotlight rendering information.
     */
    RenderListDataOptions[RenderListDataOptions["Spotlight"] = 256] = "Spotlight";
    /**
     * Returns Visualization rendering information.
     */
    RenderListDataOptions[RenderListDataOptions["Visualization"] = 512] = "Visualization";
    /**
     * Returns view XML and other information about the current view.
     */
    RenderListDataOptions[RenderListDataOptions["ViewMetadata"] = 1024] = "ViewMetadata";
    /**
     * Prevents AutoHyperlink from being run on text fields in this query.
     */
    RenderListDataOptions[RenderListDataOptions["DisableAutoHyperlink"] = 2048] = "DisableAutoHyperlink";
    /**
     * Enables urls pointing to Media TA service, such as .thumbnailUrl, .videoManifestUrl, .pdfConversionUrls.
     */
    RenderListDataOptions[RenderListDataOptions["EnableMediaTAUrls"] = 4096] = "EnableMediaTAUrls";
    /**
     * Return Parant folder information.
     */
    RenderListDataOptions[RenderListDataOptions["ParentInfo"] = 8192] = "ParentInfo";
    /**
     * Return Page context info for the current list being rendered.
     */
    RenderListDataOptions[RenderListDataOptions["PageContextInfo"] = 16384] = "PageContextInfo";
    /**
     * Return client-side component manifest information associated with the list.
     */
    RenderListDataOptions[RenderListDataOptions["ClientSideComponentManifest"] = 32768] = "ClientSideComponentManifest";
    /**
     * Return all content-types available on the list.
     */
    RenderListDataOptions[RenderListDataOptions["ListAvailableContentTypes"] = 65536] = "ListAvailableContentTypes";
    /**
      * Return the order of items in the new-item menu.
      */
    RenderListDataOptions[RenderListDataOptions["FolderContentTypeOrder"] = 131072] = "FolderContentTypeOrder";
    /**
     * Return information to initialize Grid for quick edit.
     */
    RenderListDataOptions[RenderListDataOptions["GridInitInfo"] = 262144] = "GridInitInfo";
    /**
     * Indicator if the vroom API of the SPItemUrl returned in MediaTAUrlGenerator should use site url as host.
     */
    RenderListDataOptions[RenderListDataOptions["SiteUrlAsMediaTASPItemHost"] = 524288] = "SiteUrlAsMediaTASPItemHost";
    /**
     * Return the files representing mount points in the list.
     */
    RenderListDataOptions[RenderListDataOptions["AddToOneDrive"] = 1048576] = "AddToOneDrive";
    /**
     * Return SPFX CustomAction.
     */
    RenderListDataOptions[RenderListDataOptions["SPFXCustomActions"] = 2097152] = "SPFXCustomActions";
    /**
     * Do not return non-SPFX CustomAction.
     */
    RenderListDataOptions[RenderListDataOptions["CustomActions"] = 4194304] = "CustomActions";
})(RenderListDataOptions || (RenderListDataOptions = {}));
/**
 * Determines the display mode of the given control or view
 */
var ControlMode;
(function (ControlMode) {
    ControlMode[ControlMode["Display"] = 1] = "Display";
    ControlMode[ControlMode["Edit"] = 2] = "Edit";
    ControlMode[ControlMode["New"] = 3] = "New";
})(ControlMode || (ControlMode = {}));


/***/ }),

/***/ 2519:
/*!*******************************************!*\
  !*** ./node_modules/@pnp/sp/lists/web.js ***!
  \*******************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 9706);
/* harmony import */ var _utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../utils/odata-url-from.js */ 905);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../utils/encode-path-str.js */ 4729);






(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web, "lists", _types_js__WEBPACK_IMPORTED_MODULE_2__.Lists);
(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web, "siteUserInfoList", _types_js__WEBPACK_IMPORTED_MODULE_2__.List);
(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web, "defaultDocumentLibrary", _types_js__WEBPACK_IMPORTED_MODULE_2__.List);
(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web, "customListTemplates", _spqueryable_js__WEBPACK_IMPORTED_MODULE_4__.SPCollection, "getcustomlisttemplates");
_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web.prototype.getList = function (listRelativeUrl) {
    return (0,_types_js__WEBPACK_IMPORTED_MODULE_2__.List)(this, `getList('${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_5__.encodePath)(listRelativeUrl)}')`);
};
_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web.prototype.getCatalog = async function (type) {
    const data = await (0,_webs_types_js__WEBPACK_IMPORTED_MODULE_1__.Web)(this, `getcatalog(${type})`).select("Id")();
    return (0,_types_js__WEBPACK_IMPORTED_MODULE_2__.List)([this, (0,_utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_3__.odataUrlFrom)(data)]);
};


/***/ }),

/***/ 7555:
/*!**************************************************!*\
  !*** ./node_modules/@pnp/sp/navigation/index.js ***!
  \**************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _fi_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../fi.js */ 6553);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 7002);
/* harmony import */ var _web_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./web.js */ 4542);




Reflect.defineProperty(_fi_js__WEBPACK_IMPORTED_MODULE_0__.SPFI.prototype, "navigation", {
    configurable: true,
    enumerable: true,
    get: function () {
        return this.create(_types_js__WEBPACK_IMPORTED_MODULE_1__.NavigationService);
    },
});


/***/ }),

/***/ 7002:
/*!**************************************************!*\
  !*** ./node_modules/@pnp/sp/navigation/types.js ***!
  \**************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Navigation: () => (/* binding */ Navigation),
/* harmony export */   NavigationService: () => (/* binding */ NavigationService)
/* harmony export */ });
/* unused harmony exports _NavigationNodes, NavigationNodes, _NavigationNode, NavigationNode, _Navigation, _NavigationService */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../decorators.js */ 6540);
/* harmony import */ var _index_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../index.js */ 7881);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/core */ 1971);






/**
 * Represents a collection of navigation nodes
 *
 */
class _NavigationNodes extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPCollection {
    /**
     * Gets a navigation node by id
     *
     * @param id The id of the node
     */
    getById(id) {
        return NavigationNode(this).concat(`(${id})`);
    }
    /**
     * Adds a new node to the collection
     *
     * @param title Display name of the node
     * @param url The url of the node
     * @param visible If true the node is visible, otherwise it is hidden (default: true)
     */
    async add(title, url, visible = true) {
        const postBody = (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)({
            IsVisible: visible,
            Title: title,
            Url: url,
        });
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(NavigationNodes(this, null), postBody);
    }
    /**
     * Moves a node to be after another node in the navigation
     *
     * @param nodeId Id of the node to move
     * @param previousNodeId Id of the node after which we move the node specified by nodeId
     */
    moveAfter(nodeId, previousNodeId) {
        const postBody = (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)({
            nodeId: nodeId,
            previousNodeId: previousNodeId,
        });
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(NavigationNodes(this, "MoveAfter"), postBody);
    }
}
const NavigationNodes = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_NavigationNodes);
/**
 * Represents an instance of a navigation node
 *
 */
class _NavigationNode extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPInstance {
    constructor() {
        super(...arguments);
        this.delete = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.deleteable)();
    }
    /**
     * Represents the child nodes of this node
     */
    get children() {
        return NavigationNodes(this, "children");
    }
    /**
     * Updates this node
     *
     * @param properties Properties used to update this node
     */
    async update(properties) {
        const data = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPostMerge)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)(properties));
        return {
            data,
            node: this,
        };
    }
}
const NavigationNode = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_NavigationNode);
/**
 * Exposes the navigation components
 *
 */
let _Navigation = class _Navigation extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPQueryable {
    /**
     * Gets the quicklaunch navigation nodes for the current context
     *
     */
    get quicklaunch() {
        return NavigationNodes(this, "quicklaunch");
    }
    /**
     * Gets the top bar navigation nodes for the current context
     *
     */
    get topNavigationBar() {
        return NavigationNodes(this, "topnavigationbar");
    }
};
_Navigation = (0,tslib__WEBPACK_IMPORTED_MODULE_4__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_5__.defaultPath)("navigation")
], _Navigation);

const Navigation = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_Navigation);
/**
 * Represents the top level navigation service
 */
class _NavigationService extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPQueryable {
    constructor(base = null, path) {
        super(base, path);
        this._url = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_3__.combine)((0,_index_js__WEBPACK_IMPORTED_MODULE_2__.extractWebUrl)(this._url), "_api/navigation", path);
    }
    /**
     * The MenuState service operation returns a Menu-State (dump) of a SiteMapProvider on a site.
     *
     * @param menuNodeKey MenuNode.Key of the start node within the SiteMapProvider If no key is provided the SiteMapProvider.RootNode will be the root of the menu state.
     * @param depth Depth of the dump. If no value is provided a dump with the depth of 10 is returned
     * @param mapProviderName The name identifying the SiteMapProvider to be used
     * @param customProperties comma seperated list of custom properties to be returned.
     */
    getMenuState(menuNodeKey = null, depth = 10, mapProviderName = null, customProperties = null) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(NavigationService(this, "MenuState"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)({
            customProperties,
            depth,
            mapProviderName,
            menuNodeKey,
        }));
    }
    /**
     * Tries to get a SiteMapNode.Key for a given URL within a site collection.
     *
     * @param currentUrl A url representing the SiteMapNode
     * @param mapProviderName The name identifying the SiteMapProvider to be used
     */
    getMenuNodeKey(currentUrl, mapProviderName = null) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(NavigationService(this, "MenuNodeKey"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)({
            currentUrl,
            mapProviderName,
        }));
    }
}
const NavigationService = (base, path) => new _NavigationService(base, path);


/***/ }),

/***/ 4542:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/navigation/web.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 7002);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web, "navigation", _types_js__WEBPACK_IMPORTED_MODULE_2__.Navigation);


/***/ }),

/***/ 6935:
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/presets/all.js ***!
  \*********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   SPCollection: () => (/* reexport safe */ _index_js__WEBPACK_IMPORTED_MODULE_35__.SPCollection),
/* harmony export */   SPFx: () => (/* reexport safe */ _index_js__WEBPACK_IMPORTED_MODULE_35__.SPFx),
/* harmony export */   spfi: () => (/* reexport safe */ _index_js__WEBPACK_IMPORTED_MODULE_35__.spfi)
/* harmony export */ });
/* harmony import */ var _appcatalog_index_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../appcatalog/index.js */ 7858);
/* harmony import */ var _attachments_index_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../attachments/index.js */ 2235);
/* harmony import */ var _clientside_pages_index_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../clientside-pages/index.js */ 5156);
/* harmony import */ var _column_defaults_index_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../column-defaults/index.js */ 4763);
/* harmony import */ var _comments_index_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../comments/index.js */ 5429);
/* harmony import */ var _content_types_index_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../content-types/index.js */ 6959);
/* harmony import */ var _features_index_js__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../features/index.js */ 9382);
/* harmony import */ var _fields_index_js__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../fields/index.js */ 3032);
/* harmony import */ var _files_index_js__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../files/index.js */ 7901);
/* harmony import */ var _folders_index_js__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ../folders/index.js */ 6267);
/* harmony import */ var _forms_index_js__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ../forms/index.js */ 8744);
/* harmony import */ var _hubsites_index_js__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! ../hubsites/index.js */ 5998);
/* harmony import */ var _items_index_js__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(/*! ../items/index.js */ 9721);
/* harmony import */ var _lists_index_js__WEBPACK_IMPORTED_MODULE_13__ = __webpack_require__(/*! ../lists/index.js */ 2033);
/* harmony import */ var _navigation_index_js__WEBPACK_IMPORTED_MODULE_14__ = __webpack_require__(/*! ../navigation/index.js */ 7555);
/* harmony import */ var _profiles_index_js__WEBPACK_IMPORTED_MODULE_15__ = __webpack_require__(/*! ../profiles/index.js */ 205);
/* harmony import */ var _publishing_sitepageservice_index_js__WEBPACK_IMPORTED_MODULE_16__ = __webpack_require__(/*! ../publishing-sitepageservice/index.js */ 9850);
/* harmony import */ var _regional_settings_index_js__WEBPACK_IMPORTED_MODULE_17__ = __webpack_require__(/*! ../regional-settings/index.js */ 7524);
/* harmony import */ var _related_items_index_js__WEBPACK_IMPORTED_MODULE_18__ = __webpack_require__(/*! ../related-items/index.js */ 7062);
/* harmony import */ var _search_index_js__WEBPACK_IMPORTED_MODULE_19__ = __webpack_require__(/*! ../search/index.js */ 7440);
/* harmony import */ var _security_index_js__WEBPACK_IMPORTED_MODULE_20__ = __webpack_require__(/*! ../security/index.js */ 1304);
/* harmony import */ var _sharing_index_js__WEBPACK_IMPORTED_MODULE_21__ = __webpack_require__(/*! ../sharing/index.js */ 8672);
/* harmony import */ var _site_designs_index_js__WEBPACK_IMPORTED_MODULE_22__ = __webpack_require__(/*! ../site-designs/index.js */ 4678);
/* harmony import */ var _site_groups_index_js__WEBPACK_IMPORTED_MODULE_23__ = __webpack_require__(/*! ../site-groups/index.js */ 2063);
/* harmony import */ var _site_scripts_index_js__WEBPACK_IMPORTED_MODULE_24__ = __webpack_require__(/*! ../site-scripts/index.js */ 4792);
/* harmony import */ var _site_users_index_js__WEBPACK_IMPORTED_MODULE_25__ = __webpack_require__(/*! ../site-users/index.js */ 1157);
/* harmony import */ var _sites_index_js__WEBPACK_IMPORTED_MODULE_26__ = __webpack_require__(/*! ../sites/index.js */ 5215);
/* harmony import */ var _social_index_js__WEBPACK_IMPORTED_MODULE_27__ = __webpack_require__(/*! ../social/index.js */ 4521);
/* harmony import */ var _sputilities_index_js__WEBPACK_IMPORTED_MODULE_28__ = __webpack_require__(/*! ../sputilities/index.js */ 4846);
/* harmony import */ var _subscriptions_index_js__WEBPACK_IMPORTED_MODULE_29__ = __webpack_require__(/*! ../subscriptions/index.js */ 6716);
/* harmony import */ var _user_custom_actions_index_js__WEBPACK_IMPORTED_MODULE_30__ = __webpack_require__(/*! ../user-custom-actions/index.js */ 4655);
/* harmony import */ var _views_index_js__WEBPACK_IMPORTED_MODULE_31__ = __webpack_require__(/*! ../views/index.js */ 459);
/* harmony import */ var _webparts_index_js__WEBPACK_IMPORTED_MODULE_32__ = __webpack_require__(/*! ../webparts/index.js */ 7765);
/* harmony import */ var _webs_index_js__WEBPACK_IMPORTED_MODULE_33__ = __webpack_require__(/*! ../webs/index.js */ 3867);
/* harmony import */ var _utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_34__ = __webpack_require__(/*! ../utils/odata-url-from.js */ 905);
/* harmony import */ var _index_js__WEBPACK_IMPORTED_MODULE_35__ = __webpack_require__(/*! ../index.js */ 7881);








































































/***/ }),

/***/ 205:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/profiles/index.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _fi_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../fi.js */ 6553);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 4759);



Reflect.defineProperty(_fi_js__WEBPACK_IMPORTED_MODULE_0__.SPFI.prototype, "profiles", {
    configurable: true,
    enumerable: true,
    get: function () {
        return this.create(_types_js__WEBPACK_IMPORTED_MODULE_1__.Profiles);
    },
});


/***/ }),

/***/ 4759:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/profiles/types.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Profiles: () => (/* binding */ Profiles)
/* harmony export */ });
/* unused harmony exports _Profiles, UrlZone */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../decorators.js */ 6540);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/core */ 1971);





class _Profiles extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPInstance {
    /**
     * Creates a new instance of the UserProfileQuery class
     *
     * @param baseUrl The url or SharePointQueryable which forms the parent of this user profile query
     */
    constructor(baseUrl, path = "_api/sp.userprofiles.peoplemanager") {
        super(baseUrl, path);
        this.clientPeoplePickerQuery = (new ClientPeoplePickerQuery(baseUrl)).using((0,_pnp_core__WEBPACK_IMPORTED_MODULE_2__.AssignFrom)(this));
        this.profileLoader = (new ProfileLoader(baseUrl)).using((0,_pnp_core__WEBPACK_IMPORTED_MODULE_2__.AssignFrom)(this));
    }
    /**
     * The url of the edit profile page for the current user
     */
    getEditProfileLink() {
        return Profiles(this, "EditProfileLink")();
    }
    /**
     * A boolean value that indicates whether the current user's "People I'm Following" list is public
     */
    getIsMyPeopleListPublic() {
        return Profiles(this, "IsMyPeopleListPublic")();
    }
    /**
     * A boolean value that indicates whether the current user is being followed by the specified user
     *
     * @param loginName The account name of the user
     */
    amIFollowedBy(loginName) {
        const q = Profiles(this, "amifollowedby(@v)");
        q.query.set("@v", `'${loginName}'`);
        return q();
    }
    /**
     * A boolean value that indicates whether the current user is following the specified user
     *
     * @param loginName The account name of the user
     */
    amIFollowing(loginName) {
        const q = Profiles(this, "amifollowing(@v)");
        q.query.set("@v", `'${loginName}'`);
        return q();
    }
    /**
     * Gets tags that the current user is following
     *
     * @param maxCount The maximum number of tags to retrieve (default is 20)
     */
    getFollowedTags(maxCount = 20) {
        return Profiles(this, `getfollowedtags(${maxCount})`)();
    }
    /**
     * Gets the people who are following the specified user
     *
     * @param loginName The account name of the user
     */
    getFollowersFor(loginName) {
        const q = Profiles(this, "getfollowersfor(@v)");
        q.query.set("@v", `'${loginName}'`);
        return q();
    }
    /**
     * Gets the people who are following the current user
     *
     */
    get myFollowers() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.SPCollection)(this, "getmyfollowers");
    }
    /**
     * Gets user properties for the current user
     *
     */
    get myProperties() {
        return Profiles(this, "getmyproperties");
    }
    /**
     * Gets the people who the specified user is following
     *
     * @param loginName The account name of the user.
     */
    getPeopleFollowedBy(loginName) {
        const q = Profiles(this, "getpeoplefollowedby(@v)");
        q.query.set("@v", `'${loginName}'`);
        return q();
    }
    /**
     * Gets user properties for the specified user.
     *
     * @param loginName The account name of the user.
     */
    getPropertiesFor(loginName) {
        const q = Profiles(this, "getpropertiesfor(@v)");
        q.query.set("@v", `'${loginName}'`);
        return q();
    }
    /**
     * Gets the 20 most popular hash tags over the past week, sorted so that the most popular tag appears first
     *
     */
    get trendingTags() {
        const q = Profiles(this, null);
        q.concat(".gettrendingtags");
        return q();
    }
    /**
     * Gets the specified user profile property for the specified user
     *
     * @param loginName The account name of the user
     * @param propertyName The case-sensitive name of the property to get
     */
    getUserProfilePropertyFor(loginName, propertyName) {
        const q = Profiles(this, `getuserprofilepropertyfor(accountname=@v, propertyname='${propertyName}')`);
        q.query.set("@v", `'${loginName}'`);
        return q();
    }
    /**
     * Removes the specified user from the user's list of suggested people to follow
     *
     * @param loginName The account name of the user
     */
    hideSuggestion(loginName) {
        const q = Profiles(this, "hidesuggestion(@v)");
        q.query.set("@v", `'${loginName}'`);
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(q);
    }
    /**
     * A boolean values that indicates whether the first user is following the second user
     *
     * @param follower The account name of the user who might be following the followee
     * @param followee The account name of the user who might be followed by the follower
     */
    isFollowing(follower, followee) {
        const q = Profiles(this, null);
        q.concat(".isfollowing(possiblefolloweraccountname=@v, possiblefolloweeaccountname=@y)");
        q.query.set("@v", `'${follower}'`);
        q.query.set("@y", `'${followee}'`);
        return q();
    }
    /**
     * Uploads and sets the user profile picture (Users can upload a picture to their own profile only). Not supported for batching.
     *
     * @param profilePicSource Blob data representing the user's picture in BMP, JPEG, or PNG format of up to 4.76MB
     */
    setMyProfilePic(profilePicSource) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = async (e) => {
                const buffer = e.target.result;
                try {
                    await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(Profiles(this, "setmyprofilepicture"), { body: buffer });
                    resolve();
                }
                catch (e) {
                    reject(e);
                }
            };
            reader.readAsArrayBuffer(profilePicSource);
        });
    }
    /**
     * Sets single value User Profile property
     *
     * @param accountName The account name of the user
     * @param propertyName Property name
     * @param propertyValue Property value
     */
    setSingleValueProfileProperty(accountName, propertyName, propertyValue) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(Profiles(this, "SetSingleValueProfileProperty"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)({
            accountName,
            propertyName,
            propertyValue,
        }));
    }
    /**
     * Sets multi valued User Profile property
     *
     * @param accountName The account name of the user
     * @param propertyName Property name
     * @param propertyValues Property values
     */
    setMultiValuedProfileProperty(accountName, propertyName, propertyValues) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(Profiles(this, "SetMultiValuedProfileProperty"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)({
            accountName,
            propertyName,
            propertyValues,
        }));
    }
    /**
     * Provisions one or more users' personal sites. (My Site administrator on SharePoint Online only)
     *
     * @param emails The email addresses of the users to provision sites for
     */
    createPersonalSiteEnqueueBulk(...emails) {
        return this.profileLoader.createPersonalSiteEnqueueBulk(emails);
    }
    /**
     * Gets the user profile of the site owner
     *
     */
    get ownerUserProfile() {
        return this.profileLoader.ownerUserProfile;
    }
    /**
     * Gets the user profile for the current user
     */
    get userProfile() {
        return this.profileLoader.userProfile;
    }
    /**
     * Enqueues creating a personal site for this user, which can be used to share documents, web pages, and other files
     *
     * @param interactiveRequest true if interactively (web) initiated request, or false (default) if non-interactively (client) initiated request
     */
    createPersonalSite(interactiveRequest = false) {
        return this.profileLoader.createPersonalSite(interactiveRequest);
    }
    /**
     * Sets the privacy settings for this profile
     *
     * @param share true to make all social data public; false to make all social data private
     */
    shareAllSocialData(share) {
        return this.profileLoader.shareAllSocialData(share);
    }
    /**
     * Resolves user or group using specified query parameters
     *
     * @param queryParams The query parameters used to perform resolve
     */
    clientPeoplePickerResolveUser(queryParams) {
        return this.clientPeoplePickerQuery.clientPeoplePickerResolveUser(queryParams);
    }
    /**
     * Searches for users or groups using specified query parameters
     *
     * @param queryParams The query parameters used to perform search
     */
    clientPeoplePickerSearchUser(queryParams) {
        return this.clientPeoplePickerQuery.clientPeoplePickerSearchUser(queryParams);
    }
}
const Profiles = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_Profiles);
let ProfileLoader = class ProfileLoader extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPQueryable {
    /**
     * Provisions one or more users' personal sites. (My Site administrator on SharePoint Online only) Doesn't support batching
     *
     * @param emails The email addresses of the users to provision sites for
     */
    createPersonalSiteEnqueueBulk(emails) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(ProfileLoaderFactory(this, "createpersonalsiteenqueuebulk"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)({ "emailIDs": emails }));
    }
    /**
     * Gets the user profile of the site owner.
     *
     */
    get ownerUserProfile() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(this.getParent(ProfileLoaderFactory, "_api/sp.userprofiles.profileloader.getowneruserprofile"));
    }
    /**
     * Gets the user profile of the current user.
     *
     */
    get userProfile() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(ProfileLoaderFactory(this, "getuserprofile"));
    }
    /**
     * Enqueues creating a personal site for this user, which can be used to share documents, web pages, and other files.
     *
     * @param interactiveRequest true if interactively (web) initiated request, or false (default) if non-interactively (client) initiated request
     */
    createPersonalSite(interactiveRequest = false) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(ProfileLoaderFactory(this, `getuserprofile/createpersonalsiteenque(${interactiveRequest})`));
    }
    /**
     * Sets the privacy settings for this profile
     *
     * @param share true to make all social data public; false to make all social data private.
     */
    shareAllSocialData(share) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(ProfileLoaderFactory(this, `getuserprofile/shareallsocialdata(${share})`));
    }
};
ProfileLoader = (0,tslib__WEBPACK_IMPORTED_MODULE_3__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_4__.defaultPath)("_api/sp.userprofiles.profileloader.getprofileloader")
], ProfileLoader);
const ProfileLoaderFactory = (baseUrl, path) => {
    return new ProfileLoader(baseUrl, path);
};
let ClientPeoplePickerQuery = class ClientPeoplePickerQuery extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPQueryable {
    /**
     * Resolves user or group using specified query parameters
     *
     * @param queryParams The query parameters used to perform resolve
     */
    async clientPeoplePickerResolveUser(queryParams) {
        const q = ClientPeoplePickerFactory(this, null);
        q.concat(".clientpeoplepickerresolveuser");
        const res = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(q, this.getBodyFrom(queryParams));
        return JSON.parse(typeof res === "object" ? res.ClientPeoplePickerResolveUser : res);
    }
    /**
     * Searches for users or groups using specified query parameters
     *
     * @param queryParams The query parameters used to perform search
     */
    async clientPeoplePickerSearchUser(queryParams) {
        const q = ClientPeoplePickerFactory(this, null);
        q.concat(".clientpeoplepickersearchuser");
        const res = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(q, this.getBodyFrom(queryParams));
        return JSON.parse(typeof res === "object" ? res.ClientPeoplePickerSearchUser : res);
    }
    /**
     * Creates ClientPeoplePickerQueryParameters request body
     *
     * @param queryParams The query parameters to create request body
     */
    getBodyFrom(queryParams) {
        return (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)({ queryParams });
    }
};
ClientPeoplePickerQuery = (0,tslib__WEBPACK_IMPORTED_MODULE_3__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_4__.defaultPath)("_api/sp.ui.applicationpages.clientpeoplepickerwebserviceinterface")
], ClientPeoplePickerQuery);
const ClientPeoplePickerFactory = (baseUrl, path) => {
    return new ClientPeoplePickerQuery(baseUrl, path);
};
/**
 * Specifies the originating zone of a request received.
 */
var UrlZone;
(function (UrlZone) {
    /**
     * Specifies the default zone used for requests unless another zone is specified.
     */
    UrlZone[UrlZone["DefaultZone"] = 0] = "DefaultZone";
    /**
     * Specifies an intranet zone.
     */
    UrlZone[UrlZone["Intranet"] = 1] = "Intranet";
    /**
     * Specifies an Internet zone.
     */
    UrlZone[UrlZone["Internet"] = 2] = "Internet";
    /**
     * Specifies a custom zone.
     */
    UrlZone[UrlZone["Custom"] = 3] = "Custom";
    /**
     * Specifies an extranet zone.
     */
    UrlZone[UrlZone["Extranet"] = 4] = "Extranet";
})(UrlZone || (UrlZone = {}));


/***/ }),

/***/ 9850:
/*!******************************************************************!*\
  !*** ./node_modules/@pnp/sp/publishing-sitepageservice/index.js ***!
  \******************************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _fi_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../fi.js */ 6553);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 6723);



Reflect.defineProperty(_fi_js__WEBPACK_IMPORTED_MODULE_0__.SPFI.prototype, "publishingSitePageService", {
    get: function () {
        return this.create(_types_js__WEBPACK_IMPORTED_MODULE_1__.SitePageService);
    },
});


/***/ }),

/***/ 6723:
/*!******************************************************************!*\
  !*** ./node_modules/@pnp/sp/publishing-sitepageservice/types.js ***!
  \******************************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   SitePageService: () => (/* binding */ SitePageService)
/* harmony export */ });
/* unused harmony export _SitePageService */
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);

class _SitePageService extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPInstance {
    constructor(baseUrl, path = "_api/SP.Publishing.SitePageService") {
        super(baseUrl, path);
    }
    /**
    * Gets current user unified group memberships
    */
    getCurrentUserMemberships() {
        const q = SitePageService(this, null);
        q.concat(".GetCurrentUserMemberships");
        return q();
    }
}
const SitePageService = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_SitePageService);


/***/ }),

/***/ 217:
/*!****************************************************************!*\
  !*** ./node_modules/@pnp/sp/regional-settings/content-type.js ***!
  \****************************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _content_types_types_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../content-types/types.js */ 8203);
/* harmony import */ var _funcs_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./funcs.js */ 5972);


_content_types_types_js__WEBPACK_IMPORTED_MODULE_0__._ContentType.prototype.titleResource = (0,_funcs_js__WEBPACK_IMPORTED_MODULE_1__.getValueForUICultureBinder)("nameResource");
_content_types_types_js__WEBPACK_IMPORTED_MODULE_0__._ContentType.prototype.descriptionResource = (0,_funcs_js__WEBPACK_IMPORTED_MODULE_1__.getValueForUICultureBinder)("descriptionResource");


/***/ }),

/***/ 8269:
/*!*********************************************************!*\
  !*** ./node_modules/@pnp/sp/regional-settings/field.js ***!
  \*********************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _fields_types_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../fields/types.js */ 5472);
/* harmony import */ var _funcs_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./funcs.js */ 5972);


_fields_types_js__WEBPACK_IMPORTED_MODULE_0__._Field.prototype.titleResource = (0,_funcs_js__WEBPACK_IMPORTED_MODULE_1__.getValueForUICultureBinder)("titleResource");
_fields_types_js__WEBPACK_IMPORTED_MODULE_0__._Field.prototype.descriptionResource = (0,_funcs_js__WEBPACK_IMPORTED_MODULE_1__.getValueForUICultureBinder)("descriptionResource");


/***/ }),

/***/ 5972:
/*!*********************************************************!*\
  !*** ./node_modules/@pnp/sp/regional-settings/funcs.js ***!
  \*********************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   getValueForUICultureBinder: () => (/* binding */ getValueForUICultureBinder)
/* harmony export */ });
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/queryable */ 6844);


function getValueForUICultureBinder(propName) {
    return function (cultureName) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)((0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.SPQueryable)(this, `${propName}/getValueForUICulture`), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)({ cultureName }));
    };
}


/***/ }),

/***/ 7524:
/*!*********************************************************!*\
  !*** ./node_modules/@pnp/sp/regional-settings/index.js ***!
  \*********************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _web_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web.js */ 6303);
/* harmony import */ var _user_custom_actions_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./user-custom-actions.js */ 7214);
/* harmony import */ var _list_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./list.js */ 7031);
/* harmony import */ var _field_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./field.js */ 8269);
/* harmony import */ var _content_type_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./content-type.js */ 217);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./types.js */ 9895);








/***/ }),

/***/ 7031:
/*!********************************************************!*\
  !*** ./node_modules/@pnp/sp/regional-settings/list.js ***!
  \********************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _lists_types_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../lists/types.js */ 9706);
/* harmony import */ var _funcs_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./funcs.js */ 5972);


_lists_types_js__WEBPACK_IMPORTED_MODULE_0__._List.prototype.titleResource = (0,_funcs_js__WEBPACK_IMPORTED_MODULE_1__.getValueForUICultureBinder)("titleResource");
_lists_types_js__WEBPACK_IMPORTED_MODULE_0__._List.prototype.descriptionResource = (0,_funcs_js__WEBPACK_IMPORTED_MODULE_1__.getValueForUICultureBinder)("descriptionResource");


/***/ }),

/***/ 9895:
/*!*********************************************************!*\
  !*** ./node_modules/@pnp/sp/regional-settings/types.js ***!
  \*********************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   RegionalSettings: () => (/* binding */ RegionalSettings)
/* harmony export */ });
/* unused harmony exports _RegionalSettings, _TimeZone, TimeZone, _TimeZones, TimeZones */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../decorators.js */ 6540);




let _RegionalSettings = class _RegionalSettings extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPInstance {
    /**
     * Gets time zone
     */
    get timeZone() {
        return TimeZone(this);
    }
    /**
     * Gets time zones
     */
    get timeZones() {
        return TimeZones(this);
    }
    /**
     * Gets the collection of languages used in a server farm.
     */
    async getInstalledLanguages() {
        const results = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.SPCollection)(this, "installedlanguages")();
        return results.Items;
    }
};
_RegionalSettings = (0,tslib__WEBPACK_IMPORTED_MODULE_2__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_3__.defaultPath)("regionalsettings")
], _RegionalSettings);

const RegionalSettings = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spInvokableFactory)(_RegionalSettings);
let _TimeZone = class _TimeZone extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPInstance {
    /**
     * Gets an Local Time by UTC Time
     *
     * @param utcTime UTC Time as Date or ISO String
     */
    async utcToLocalTime(utcTime) {
        let dateIsoString;
        if (typeof utcTime === "string") {
            dateIsoString = utcTime;
        }
        else {
            dateIsoString = utcTime.toISOString();
        }
        const res = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(TimeZone(this, `utctolocaltime('${dateIsoString}')`));
        return (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(res, "UTCToLocalTime") ? res.UTCToLocalTime : res;
    }
    /**
     * Gets an UTC Time by Local Time
     *
     * @param localTime Local Time as Date or ISO String
     */
    async localTimeToUTC(localTime) {
        let dateIsoString;
        if (typeof localTime === "string") {
            dateIsoString = localTime;
        }
        else {
            dateIsoString = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.dateAdd)(localTime, "minute", localTime.getTimezoneOffset() * -1).toISOString();
        }
        const res = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(TimeZone(this, `localtimetoutc('${dateIsoString}')`));
        return (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(res, "LocalTimeToUTC") ? res.LocalTimeToUTC : res;
    }
};
_TimeZone = (0,tslib__WEBPACK_IMPORTED_MODULE_2__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_3__.defaultPath)("timezone")
], _TimeZone);

const TimeZone = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spInvokableFactory)(_TimeZone);
let _TimeZones = class _TimeZones extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPCollection {
    /**
     * Gets an TimeZone by id (see: https://msdn.microsoft.com/en-us/library/office/jj247008.aspx)
     *
     * @param id The integer id of the timezone to retrieve
     */
    getById(id) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(TimeZones(this, `GetById(${id})`));
    }
};
_TimeZones = (0,tslib__WEBPACK_IMPORTED_MODULE_2__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_3__.defaultPath)("timezones")
], _TimeZones);

const TimeZones = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spInvokableFactory)(_TimeZones);


/***/ }),

/***/ 7214:
/*!***********************************************************************!*\
  !*** ./node_modules/@pnp/sp/regional-settings/user-custom-actions.js ***!
  \***********************************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _user_custom_actions_types_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../user-custom-actions/types.js */ 6687);
/* harmony import */ var _funcs_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./funcs.js */ 5972);


_user_custom_actions_types_js__WEBPACK_IMPORTED_MODULE_0__._UserCustomAction.prototype.titleResource = (0,_funcs_js__WEBPACK_IMPORTED_MODULE_1__.getValueForUICultureBinder)("titleResource");
_user_custom_actions_types_js__WEBPACK_IMPORTED_MODULE_0__._UserCustomAction.prototype.descriptionResource = (0,_funcs_js__WEBPACK_IMPORTED_MODULE_1__.getValueForUICultureBinder)("descriptionResource");


/***/ }),

/***/ 6303:
/*!*******************************************************!*\
  !*** ./node_modules/@pnp/sp/regional-settings/web.js ***!
  \*******************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 9895);
/* harmony import */ var _funcs_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./funcs.js */ 5972);




(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web, "regionalSettings", _types_js__WEBPACK_IMPORTED_MODULE_2__.RegionalSettings);
_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web.prototype.titleResource = (0,_funcs_js__WEBPACK_IMPORTED_MODULE_3__.getValueForUICultureBinder)("titleResource");
_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web.prototype.descriptionResource = (0,_funcs_js__WEBPACK_IMPORTED_MODULE_3__.getValueForUICultureBinder)("descriptionResource");


/***/ }),

/***/ 7062:
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/sp/related-items/index.js ***!
  \*****************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _web_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web.js */ 6368);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 698);




/***/ }),

/***/ 698:
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/sp/related-items/types.js ***!
  \*****************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   RelatedItemManager: () => (/* binding */ RelatedItemManager)
/* harmony export */ });
/* unused harmony export _RelatedItemManager */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../utils/extract-web-url.js */ 2997);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../decorators.js */ 6540);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/queryable */ 6844);





let _RelatedItemManager = class _RelatedItemManager extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPQueryable {
    getRelatedItems(sourceListName, sourceItemId) {
        const query = RelatedItemManager(this);
        query.concat(".GetRelatedItems");
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(query, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_2__.body)({
            SourceItemID: sourceItemId,
            SourceListName: sourceListName,
        }));
    }
    getPageOneRelatedItems(sourceListName, sourceItemId) {
        const query = RelatedItemManager(this);
        query.concat(".GetPageOneRelatedItems");
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(query, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_2__.body)({
            SourceItemID: sourceItemId,
            SourceListName: sourceListName,
        }));
    }
    addSingleLink(sourceListName, sourceItemId, sourceWebUrl, targetListName, targetItemID, targetWebUrl, tryAddReverseLink = false) {
        const query = RelatedItemManager(this);
        query.concat(".AddSingleLink");
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(query, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_2__.body)({
            SourceItemID: sourceItemId,
            SourceListName: sourceListName,
            SourceWebUrl: sourceWebUrl,
            TargetItemID: targetItemID,
            TargetListName: targetListName,
            TargetWebUrl: targetWebUrl,
            TryAddReverseLink: tryAddReverseLink,
        }));
    }
    addSingleLinkToUrl(sourceListName, sourceItemId, targetItemUrl, tryAddReverseLink = false) {
        const query = RelatedItemManager(this);
        query.concat(".AddSingleLinkToUrl");
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(query, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_2__.body)({
            SourceItemID: sourceItemId,
            SourceListName: sourceListName,
            TargetItemUrl: targetItemUrl,
            TryAddReverseLink: tryAddReverseLink,
        }));
    }
    addSingleLinkFromUrl(sourceItemUrl, targetListName, targetItemId, tryAddReverseLink = false) {
        const query = RelatedItemManager(this);
        query.concat(".AddSingleLinkFromUrl");
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(query, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_2__.body)({
            SourceItemUrl: sourceItemUrl,
            TargetItemID: targetItemId,
            TargetListName: targetListName,
            TryAddReverseLink: tryAddReverseLink,
        }));
    }
    deleteSingleLink(sourceListName, sourceItemId, sourceWebUrl, targetListName, targetItemId, targetWebUrl, tryDeleteReverseLink = false) {
        const query = RelatedItemManager(this);
        query.concat(".DeleteSingleLink");
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(query, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_2__.body)({
            SourceItemID: sourceItemId,
            SourceListName: sourceListName,
            SourceWebUrl: sourceWebUrl,
            TargetItemID: targetItemId,
            TargetListName: targetListName,
            TargetWebUrl: targetWebUrl,
            TryDeleteReverseLink: tryDeleteReverseLink,
        }));
    }
};
_RelatedItemManager = (0,tslib__WEBPACK_IMPORTED_MODULE_3__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_4__.defaultPath)("_api/SP.RelatedItemManager")
], _RelatedItemManager);

const RelatedItemManager = (base) => {
    if (typeof base === "string") {
        return new _RelatedItemManager((0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_1__.extractWebUrl)(base));
    }
    return new _RelatedItemManager([base, (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_1__.extractWebUrl)(base.toUrl())]);
};


/***/ }),

/***/ 6368:
/*!***************************************************!*\
  !*** ./node_modules/@pnp/sp/related-items/web.js ***!
  \***************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 698);


Reflect.defineProperty(_webs_types_js__WEBPACK_IMPORTED_MODULE_0__._Web.prototype, "relatedItems", {
    configurable: true,
    enumerable: true,
    get: function () {
        return (0,_types_js__WEBPACK_IMPORTED_MODULE_1__.RelatedItemManager)(this);
    },
});


/***/ }),

/***/ 7440:
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/search/index.js ***!
  \**********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _fi_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../fi.js */ 6553);
/* harmony import */ var _query_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./query.js */ 3147);
/* harmony import */ var _suggest_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./suggest.js */ 2146);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./types.js */ 686);






_fi_js__WEBPACK_IMPORTED_MODULE_0__.SPFI.prototype.search = function (query) {
    return (new _query_js__WEBPACK_IMPORTED_MODULE_1__._Search(this._root)).run(query);
};
_fi_js__WEBPACK_IMPORTED_MODULE_0__.SPFI.prototype.searchSuggest = function (query) {
    return (new _suggest_js__WEBPACK_IMPORTED_MODULE_2__._Suggest(this._root)).run(typeof query === "string" ? { querytext: query } : query);
};


/***/ }),

/***/ 3147:
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/search/query.js ***!
  \**********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   _Search: () => (/* binding */ _Search)
/* harmony export */ });
/* unused harmony exports SearchQueryBuilder, Search, SearchResults */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../decorators.js */ 6540);
var _Search_1;





const funcs = new Map([
    ["text", "Querytext"],
    ["template", "QueryTemplate"],
    ["sourceId", "SourceId"],
    ["trimDuplicatesIncludeId", ""],
    ["startRow", ""],
    ["rowLimit", ""],
    ["rankingModelId", ""],
    ["rowsPerPage", ""],
    ["selectProperties", ""],
    ["culture", ""],
    ["timeZoneId", ""],
    ["refinementFilters", ""],
    ["refiners", ""],
    ["hiddenConstraints", ""],
    ["sortList", ""],
    ["timeout", ""],
    ["hithighlightedProperties", ""],
    ["clientType", ""],
    ["personalizationData", ""],
    ["resultsURL", ""],
    ["queryTag", ""],
    ["properties", ""],
    ["queryTemplatePropertiesUrl", ""],
    ["reorderingRules", ""],
    ["hitHighlightedMultivaluePropertyLimit", ""],
    ["collapseSpecification", ""],
    ["uiLanguage", ""],
    ["desiredSnippetLength", ""],
    ["maxSnippetLength", ""],
    ["summaryLength", ""],
]);
const props = new Map([]);
function toPropCase(str) {
    return str.replace(/^(.)/, ($1) => $1.toUpperCase());
}
/**
 * Creates a new instance of the SearchQueryBuilder
 *
 * @param queryText Initial query text
 * @param _query Any initial query configuration
 */
function SearchQueryBuilder(queryText = "", _query = {}) {
    return new Proxy({
        query: Object.assign({
            Querytext: queryText,
        }, _query),
    }, {
        get(self, propertyKey, proxy) {
            const pk = propertyKey.toString();
            if (pk === "toSearchQuery") {
                return () => self.query;
            }
            if (funcs.has(pk)) {
                return (...value) => {
                    const mappedPk = funcs.get(pk);
                    self.query[mappedPk.length > 0 ? mappedPk : toPropCase(pk)] = value.length > 1 ? value : value[0];
                    return proxy;
                };
            }
            const propKey = props.has(pk) ? props.get(pk) : toPropCase(pk);
            self.query[propKey] = true;
            return proxy;
        },
    });
}
/**
 * Describes the search API
 *
 */
let _Search = _Search_1 = class _Search extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPInstance {
    /**
     * @returns Promise
     */
    async run(queryInit) {
        const query = this.parseQuery(queryInit);
        const postBody = (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_2__.body)({
            request: {
                ...query,
                HitHighlightedProperties: this.fixArrProp(query.HitHighlightedProperties),
                Properties: this.fixArrProp(query.Properties),
                RefinementFilters: this.fixArrProp(query.RefinementFilters),
                ReorderingRules: this.fixArrProp(query.ReorderingRules),
                SelectProperties: this.fixArrProp(query.SelectProperties),
                SortList: this.fixArrProp(query.SortList),
            },
        });
        const poster = new _Search_1([this, this.parentUrl]);
        poster.using((0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_2__.CacheAlways)(), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_2__.CacheKey)((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.getHashCode)(JSON.stringify(postBody)).toString()));
        const data = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(poster, postBody);
        // Create search instance copy for SearchResult's getPage request.
        return new SearchResults(data, new _Search_1([this, this.parentUrl]), query);
    }
    /**
     * Fix array property
     *
     * @param prop property to fix for container struct
     */
    fixArrProp(prop) {
        return typeof prop === "undefined" ? [] : (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.isArray)(prop) ? prop : [prop];
    }
    /**
     * Translates one of the query initializers into a SearchQuery instance
     *
     * @param query
     */
    parseQuery(query) {
        let finalQuery;
        if (typeof query === "string") {
            finalQuery = { Querytext: query };
        }
        else if (query.toSearchQuery) {
            finalQuery = query.toSearchQuery();
        }
        else {
            finalQuery = query;
        }
        return finalQuery;
    }
};
_Search = _Search_1 = (0,tslib__WEBPACK_IMPORTED_MODULE_3__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_4__.defaultPath)("_api/search/postquery"),
    (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_2__.invokable)(function (init) {
        return this.run(init);
    })
], _Search);

const Search = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_Search);
class SearchResults {
    constructor(rawResponse, _search, _query, _raw = null, _primary = null) {
        this._search = _search;
        this._query = _query;
        this._raw = _raw;
        this._primary = _primary;
        this._raw = rawResponse.postquery ? rawResponse.postquery : rawResponse;
    }
    get ElapsedTime() {
        var _a;
        return ((_a = this === null || this === void 0 ? void 0 : this.RawSearchResults) === null || _a === void 0 ? void 0 : _a.ElapsedTime) || 0;
    }
    get RowCount() {
        var _a, _b, _c;
        return ((_c = (_b = (_a = this === null || this === void 0 ? void 0 : this.RawSearchResults) === null || _a === void 0 ? void 0 : _a.PrimaryQueryResult) === null || _b === void 0 ? void 0 : _b.RelevantResults) === null || _c === void 0 ? void 0 : _c.RowCount) || 0;
    }
    get TotalRows() {
        var _a, _b, _c;
        return ((_c = (_b = (_a = this === null || this === void 0 ? void 0 : this.RawSearchResults) === null || _a === void 0 ? void 0 : _a.PrimaryQueryResult) === null || _b === void 0 ? void 0 : _b.RelevantResults) === null || _c === void 0 ? void 0 : _c.TotalRows) || 0;
    }
    get TotalRowsIncludingDuplicates() {
        var _a, _b, _c;
        return ((_c = (_b = (_a = this === null || this === void 0 ? void 0 : this.RawSearchResults) === null || _a === void 0 ? void 0 : _a.PrimaryQueryResult) === null || _b === void 0 ? void 0 : _b.RelevantResults) === null || _c === void 0 ? void 0 : _c.TotalRowsIncludingDuplicates) || 0;
    }
    get RawSearchResults() {
        return this._raw;
    }
    get PrimarySearchResults() {
        var _a, _b, _c, _d;
        if (this._primary === null) {
            this._primary = this.formatSearchResults(((_d = (_c = (_b = (_a = this._raw) === null || _a === void 0 ? void 0 : _a.PrimaryQueryResult) === null || _b === void 0 ? void 0 : _b.RelevantResults) === null || _c === void 0 ? void 0 : _c.Table) === null || _d === void 0 ? void 0 : _d.Rows) || null);
        }
        return this._primary;
    }
    /**
     * Gets a page of results
     *
     * @param pageNumber Index of the page to return. Used to determine StartRow
     * @param pageSize Optional, items per page (default = 10)
     */
    getPage(pageNumber, pageSize) {
        // if we got all the available rows we don't have another page
        if (this.TotalRows < this.RowCount) {
            return Promise.resolve(null);
        }
        // if pageSize is supplied, then we use that regardless of any previous values
        // otherwise get the previous RowLimit or default to 10
        const rows = pageSize !== undefined ? pageSize : (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(this._query, "RowLimit") ? this._query.RowLimit : 10;
        const query = {
            ...this._query,
            RowLimit: rows,
            StartRow: rows * (pageNumber - 1),
        };
        // we have reached the end
        if (query.StartRow > this.TotalRows) {
            return Promise.resolve(null);
        }
        return this._search.run(query);
    }
    /**
     * Formats a search results array
     *
     * @param rawResults The array to process
     */
    formatSearchResults(rawResults) {
        const results = new Array();
        if (typeof (rawResults) === "undefined" || rawResults == null) {
            return [];
        }
        const tempResults = rawResults.results ? rawResults.results : rawResults;
        for (const tempResult of tempResults) {
            const cells = tempResult.Cells.results ? tempResult.Cells.results : tempResult.Cells;
            results.push(cells.reduce((res, cell) => {
                res[cell.Key] = cell.Value;
                return res;
            }, {}));
        }
        return results;
    }
}


/***/ }),

/***/ 2146:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/search/suggest.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   _Suggest: () => (/* binding */ _Suggest)
/* harmony export */ });
/* unused harmony export Suggest */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../decorators.js */ 6540);




let _Suggest = class _Suggest extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPInstance {
    async run(query) {
        this.mapQueryToQueryString(query);
        const response = await this();
        const mapper = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(response, "suggest") ? (s_1) => response.suggest[s_1].results : (s_2) => response[s_2];
        return {
            PeopleNames: mapper("PeopleNames"),
            PersonalResults: mapper("PersonalResults"),
            Queries: mapper("Queries"),
        };
    }
    mapQueryToQueryString(query) {
        const setProp = (q) => (checkProp) => (sp) => {
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(q, checkProp)) {
                this.query.set(sp, q[checkProp].toString());
            }
        };
        this.query.set("querytext", `'${query.querytext}'`);
        const querySetter = setProp(query);
        querySetter("count")("inumberofquerysuggestions");
        querySetter("personalCount")("inumberofresultsuggestions");
        querySetter("preQuery")("fprequerysuggestions");
        querySetter("hitHighlighting")("fhithighlighting");
        querySetter("capitalize")("fcapitalizefirstletters");
        querySetter("culture")("culture");
        querySetter("stemming")("enablestemming");
        querySetter("includePeople")("showpeoplenamesuggestions");
        querySetter("queryRules")("enablequeryrules");
        querySetter("prefixMatch")("fprefixmatchallterms");
    }
};
_Suggest = (0,tslib__WEBPACK_IMPORTED_MODULE_2__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_3__.defaultPath)("_api/search/suggest")
], _Suggest);

const Suggest = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_Suggest);


/***/ }),

/***/ 686:
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/search/types.js ***!
  \**********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* unused harmony exports SortDirection, ReorderingRuleMatchType, QueryPropertyValueType, SearchBuiltInSourceId */
/**
 * defines the SortDirection enum
 */
var SortDirection;
(function (SortDirection) {
    SortDirection[SortDirection["Ascending"] = 0] = "Ascending";
    SortDirection[SortDirection["Descending"] = 1] = "Descending";
    SortDirection[SortDirection["FQLFormula"] = 2] = "FQLFormula";
})(SortDirection || (SortDirection = {}));
/**
 * defines the ReorderingRuleMatchType  enum
 */
var ReorderingRuleMatchType;
(function (ReorderingRuleMatchType) {
    ReorderingRuleMatchType[ReorderingRuleMatchType["ResultContainsKeyword"] = 0] = "ResultContainsKeyword";
    ReorderingRuleMatchType[ReorderingRuleMatchType["TitleContainsKeyword"] = 1] = "TitleContainsKeyword";
    ReorderingRuleMatchType[ReorderingRuleMatchType["TitleMatchesKeyword"] = 2] = "TitleMatchesKeyword";
    ReorderingRuleMatchType[ReorderingRuleMatchType["UrlStartsWith"] = 3] = "UrlStartsWith";
    ReorderingRuleMatchType[ReorderingRuleMatchType["UrlExactlyMatches"] = 4] = "UrlExactlyMatches";
    ReorderingRuleMatchType[ReorderingRuleMatchType["ContentTypeIs"] = 5] = "ContentTypeIs";
    ReorderingRuleMatchType[ReorderingRuleMatchType["FileExtensionMatches"] = 6] = "FileExtensionMatches";
    ReorderingRuleMatchType[ReorderingRuleMatchType["ResultHasTag"] = 7] = "ResultHasTag";
    ReorderingRuleMatchType[ReorderingRuleMatchType["ManualCondition"] = 8] = "ManualCondition";
})(ReorderingRuleMatchType || (ReorderingRuleMatchType = {}));
/**
 * Specifies the type value for the property
 */
var QueryPropertyValueType;
(function (QueryPropertyValueType) {
    QueryPropertyValueType[QueryPropertyValueType["None"] = 0] = "None";
    QueryPropertyValueType[QueryPropertyValueType["StringType"] = 1] = "StringType";
    QueryPropertyValueType[QueryPropertyValueType["Int32Type"] = 2] = "Int32Type";
    QueryPropertyValueType[QueryPropertyValueType["BooleanType"] = 3] = "BooleanType";
    QueryPropertyValueType[QueryPropertyValueType["StringArrayType"] = 4] = "StringArrayType";
    QueryPropertyValueType[QueryPropertyValueType["UnSupportedType"] = 5] = "UnSupportedType";
})(QueryPropertyValueType || (QueryPropertyValueType = {}));
class SearchBuiltInSourceId {
}
SearchBuiltInSourceId.Documents = "e7ec8cee-ded8-43c9-beb5-436b54b31e84";
SearchBuiltInSourceId.ItemsMatchingContentType = "5dc9f503-801e-4ced-8a2c-5d1237132419";
SearchBuiltInSourceId.ItemsMatchingTag = "e1327b9c-2b8c-4b23-99c9-3730cb29c3f7";
SearchBuiltInSourceId.ItemsRelatedToCurrentUser = "48fec42e-4a92-48ce-8363-c2703a40e67d";
SearchBuiltInSourceId.ItemsWithSameKeywordAsThisItem = "5c069288-1d17-454a-8ac6-9c642a065f48";
SearchBuiltInSourceId.LocalPeopleResults = "b09a7990-05ea-4af9-81ef-edfab16c4e31";
SearchBuiltInSourceId.LocalReportsAndDataResults = "203fba36-2763-4060-9931-911ac8c0583b";
SearchBuiltInSourceId.LocalSharePointResults = "8413cd39-2156-4e00-b54d-11efd9abdb89";
SearchBuiltInSourceId.LocalVideoResults = "78b793ce-7956-4669-aa3b-451fc5defebf";
SearchBuiltInSourceId.Pages = "5e34578e-4d08-4edc-8bf3-002acf3cdbcc";
SearchBuiltInSourceId.Pictures = "38403c8c-3975-41a8-826e-717f2d41568a";
SearchBuiltInSourceId.Popular = "97c71db1-58ce-4891-8b64-585bc2326c12";
SearchBuiltInSourceId.RecentlyChangedItems = "ba63bbae-fa9c-42c0-b027-9a878f16557c";
SearchBuiltInSourceId.RecommendedItems = "ec675252-14fa-4fbe-84dd-8d098ed74181";
SearchBuiltInSourceId.Wiki = "9479bf85-e257-4318-b5a8-81a180f5faa1";


/***/ }),

/***/ 3974:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/security/funcs.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   breakRoleInheritance: () => (/* binding */ breakRoleInheritance),
/* harmony export */   currentUserHasPermissions: () => (/* binding */ currentUserHasPermissions),
/* harmony export */   getCurrentUserEffectivePermissions: () => (/* binding */ getCurrentUserEffectivePermissions),
/* harmony export */   getUserEffectivePermissions: () => (/* binding */ getUserEffectivePermissions),
/* harmony export */   hasPermissions: () => (/* binding */ hasPermissions),
/* harmony export */   resetRoleInheritance: () => (/* binding */ resetRoleInheritance),
/* harmony export */   userHasPermissions: () => (/* binding */ userHasPermissions)
/* harmony export */ });
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./types.js */ 9713);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../spqueryable.js */ 2678);


/**
* Gets the effective permissions for the user supplied
*
* @param loginName The claims username for the user (ex: i:0#.f|membership|user@domain.com)
*/
async function getUserEffectivePermissions(loginName) {
    const q = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.SPInstance)(this, "getUserEffectivePermissions(@user)");
    q.query.set("@user", `'${loginName}'`);
    return q();
}
/**
 * Gets the effective permissions for the current user
 */
async function getCurrentUserEffectivePermissions() {
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.SPQueryable)(this, "EffectiveBasePermissions")();
}
/**
 * Breaks the security inheritance at this level optinally copying permissions and clearing subscopes
 *
 * @param copyRoleAssignments If true the permissions are copied from the current parent scope
 * @param clearSubscopes Optional. true to make all child securable objects inherit role assignments from the current object
 */
async function breakRoleInheritance(copyRoleAssignments = false, clearSubscopes = false) {
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)((0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.SPQueryable)(this, `breakroleinheritance(copyroleassignments=${copyRoleAssignments}, clearsubscopes=${clearSubscopes})`));
}
/**
 * Removes the local role assignments so that it re-inherit role assignments from the parent object.
 *
 */
async function resetRoleInheritance() {
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)((0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.SPQueryable)(this, "resetroleinheritance"));
}
/**
 * Determines if a given user has the appropriate permissions
 *
 * @param loginName The user to check
 * @param permission The permission being checked
 */
async function userHasPermissions(loginName, permission) {
    const perms = await getUserEffectivePermissions.call(this, loginName);
    return this.hasPermissions(perms, permission);
}
/**
 * Determines if the current user has the requested permissions
 *
 * @param permission The permission we wish to check
 */
async function currentUserHasPermissions(permission) {
    const perms = await getCurrentUserEffectivePermissions.call(this);
    return this.hasPermissions(perms, permission);
}
/**
 * Taken from sp.js, checks the supplied permissions against the mask
 *
 * @param value The security principal's permissions on the given object
 * @param perm The permission checked against the value
 */
/* eslint-disable no-bitwise */
function hasPermissions(value, perm) {
    if (!perm) {
        return true;
    }
    if (perm === _types_js__WEBPACK_IMPORTED_MODULE_0__.PermissionKind.FullMask) {
        return (value.High & 32767) === 32767 && value.Low === 65535;
    }
    perm = perm - 1;
    let num = 1;
    if (perm >= 0 && perm < 32) {
        num = num << perm;
        return 0 !== (value.Low & num);
    }
    else if (perm >= 32 && perm < 64) {
        num = num << perm - 32;
        return 0 !== (value.High & num);
    }
    return false;
}
/* eslint-enable no-bitwise */


/***/ }),

/***/ 1304:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/security/index.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _item_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./item.js */ 7780);
/* harmony import */ var _list_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./list.js */ 2207);
/* harmony import */ var _web_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./web.js */ 3844);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./types.js */ 9713);






/***/ }),

/***/ 7780:
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/security/item.js ***!
  \***********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _items_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../items/types.js */ 132);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 9713);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _funcs_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./funcs.js */ 3974);





(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_items_types_js__WEBPACK_IMPORTED_MODULE_1__._Item, "roleAssignments", _types_js__WEBPACK_IMPORTED_MODULE_2__.RoleAssignments);
(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_items_types_js__WEBPACK_IMPORTED_MODULE_1__._Item, "firstUniqueAncestorSecurableObject", _spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.SPInstance);
_items_types_js__WEBPACK_IMPORTED_MODULE_1__._Item.prototype.getUserEffectivePermissions = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.getUserEffectivePermissions;
_items_types_js__WEBPACK_IMPORTED_MODULE_1__._Item.prototype.getCurrentUserEffectivePermissions = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.getCurrentUserEffectivePermissions;
_items_types_js__WEBPACK_IMPORTED_MODULE_1__._Item.prototype.breakRoleInheritance = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.breakRoleInheritance;
_items_types_js__WEBPACK_IMPORTED_MODULE_1__._Item.prototype.resetRoleInheritance = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.resetRoleInheritance;
_items_types_js__WEBPACK_IMPORTED_MODULE_1__._Item.prototype.userHasPermissions = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.userHasPermissions;
_items_types_js__WEBPACK_IMPORTED_MODULE_1__._Item.prototype.currentUserHasPermissions = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.currentUserHasPermissions;
_items_types_js__WEBPACK_IMPORTED_MODULE_1__._Item.prototype.hasPermissions = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.hasPermissions;


/***/ }),

/***/ 2207:
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/security/list.js ***!
  \***********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _lists_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types.js */ 9706);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 9713);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _funcs_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./funcs.js */ 3974);





(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List, "roleAssignments", _types_js__WEBPACK_IMPORTED_MODULE_2__.RoleAssignments);
(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List, "firstUniqueAncestorSecurableObject", _spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.SPInstance);
_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List.prototype.getUserEffectivePermissions = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.getUserEffectivePermissions;
_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List.prototype.getCurrentUserEffectivePermissions = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.getCurrentUserEffectivePermissions;
_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List.prototype.breakRoleInheritance = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.breakRoleInheritance;
_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List.prototype.resetRoleInheritance = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.resetRoleInheritance;
_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List.prototype.userHasPermissions = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.userHasPermissions;
_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List.prototype.currentUserHasPermissions = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.currentUserHasPermissions;
_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List.prototype.hasPermissions = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.hasPermissions;


/***/ }),

/***/ 9713:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/security/types.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   PermissionKind: () => (/* binding */ PermissionKind),
/* harmony export */   RoleAssignments: () => (/* binding */ RoleAssignments),
/* harmony export */   RoleDefinitions: () => (/* binding */ RoleDefinitions)
/* harmony export */ });
/* unused harmony exports _RoleAssignments, _RoleAssignment, RoleAssignment, _RoleDefinitions, _RoleDefinition, RoleDefinition */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _site_groups_types_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../site-groups/types.js */ 4281);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../decorators.js */ 6540);






/**
 * Describes a set of role assignments for the current scope
 *
 */
let _RoleAssignments = class _RoleAssignments extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_2__._SPCollection {
    /**
     * Gets the role assignment associated with the specified principal id from the collection.
     *
     * @param id The id of the role assignment
     */
    getById(id) {
        return RoleAssignment(this).concat(`(${id})`);
    }
    /**
     * Adds a new role assignment with the specified principal and role definitions to the collection
     *
     * @param principalId The id of the user or group to assign permissions to
     * @param roleDefId The id of the role definition that defines the permissions to assign
     *
     */
    async add(principalId, roleDefId) {
        await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(RoleAssignments(this, `addroleassignment(principalid=${principalId}, roledefid=${roleDefId})`));
    }
    /**
     * Removes the role assignment with the specified principal and role definition from the collection
     *
     * @param principalId The id of the user or group in the role assignment
     * @param roleDefId The id of the role definition in the role assignment
     *
     */
    async remove(principalId, roleDefId) {
        await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(RoleAssignments(this, `removeroleassignment(principalid=${principalId}, roledefid=${roleDefId})`));
    }
};
_RoleAssignments = (0,tslib__WEBPACK_IMPORTED_MODULE_4__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_5__.defaultPath)("roleassignments")
], _RoleAssignments);

const RoleAssignments = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spInvokableFactory)(_RoleAssignments);
/**
 * Describes a role assignment
 *
 */
class _RoleAssignment extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_2__._SPInstance {
    constructor() {
        super(...arguments);
        this.delete = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.deleteable)();
    }
    /**
     * Gets the groups that directly belong to the access control list (ACL) for this securable object
     *
     */
    get groups() {
        return (0,_site_groups_types_js__WEBPACK_IMPORTED_MODULE_3__.SiteGroups)(this, "groups");
    }
    /**
     * Gets the role definition bindings for this role assignment
     *
     */
    get bindings() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPCollection)(this, "roledefinitionbindings");
    }
}
const RoleAssignment = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spInvokableFactory)(_RoleAssignment);
/**
 * Describes a collection of role definitions
 *
 */
let _RoleDefinitions = class _RoleDefinitions extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_2__._SPCollection {
    /**
     * Gets the role definition with the specified id from the collection
     *
     * @param id The id of the role definition
     *
     */
    getById(id) {
        return RoleDefinition(this, `getById(${id})`);
    }
    /**
     * Gets the role definition with the specified name
     *
     * @param name The name of the role definition
     *
     */
    getByName(name) {
        return RoleDefinition(this, `getbyname('${name}')`);
    }
    /**
     * Gets the role definition with the specified role type
     *
     * @param roleTypeKind The roletypekind of the role definition (None=0, Guest=1, Reader=2, Contributor=3, WebDesigner=4, Administrator=5, Editor=6, System=7)
     *
     */
    getByType(roleTypeKind) {
        return RoleDefinition(this, `getbytype(${roleTypeKind})`);
    }
    /**
     * Creates a role definition
     *
     * @param name The new role definition's name
     * @param description The new role definition's description
     * @param order The order in which the role definition appears
     * @param basePermissions The permissions mask for this role definition, high and low values need to be converted to string
     *
     */
    async add(name, description, order, basePermissions) {
        const postBody = (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)({
            BasePermissions: { "High": basePermissions.High.toString(), "Low": basePermissions.Low.toString() },
            Description: description,
            Name: name,
            Order: order,
        });
        // __metadata: { "type": "SP.RoleDefinition" },
        const data = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(this, postBody);
        return {
            data: data,
            definition: this.getById(data.Id),
        };
    }
};
_RoleDefinitions = (0,tslib__WEBPACK_IMPORTED_MODULE_4__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_5__.defaultPath)("roledefinitions")
], _RoleDefinitions);

const RoleDefinitions = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spInvokableFactory)(_RoleDefinitions);
/**
 * Describes a role definition
 *
 */
class _RoleDefinition extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_2__._SPInstance {
    constructor() {
        super(...arguments);
        this.delete = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.deleteable)();
    }
    /**
     * Updates this role definition with the supplied properties
     *
     * @param properties A plain object hash of values to update for the role definition
     */
    async update(properties) {
        const s = ["BasePermissions"];
        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(properties, s[0]) !== undefined) {
            const bpObj = properties[s[0]];
            bpObj.High = bpObj.High.toString();
            bpObj.Low = bpObj.Low.toString();
        }
        const data = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPostMerge)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)(properties));
        let definition = this;
        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(properties, "Name")) {
            const parent = this.getParent(RoleDefinitions);
            definition = parent.getByName(properties.Name);
        }
        return {
            data,
            definition,
        };
    }
}
const RoleDefinition = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spInvokableFactory)(_RoleDefinition);
var PermissionKind;
(function (PermissionKind) {
    /**
     * Has no permissions on the Site. Not available through the user interface.
     */
    PermissionKind[PermissionKind["EmptyMask"] = 0] = "EmptyMask";
    /**
     * View items in lists, documents in document libraries, and Web discussion comments.
     */
    PermissionKind[PermissionKind["ViewListItems"] = 1] = "ViewListItems";
    /**
     * Add items to lists, documents to document libraries, and Web discussion comments.
     */
    PermissionKind[PermissionKind["AddListItems"] = 2] = "AddListItems";
    /**
     * Edit items in lists, edit documents in document libraries, edit Web discussion comments
     * in documents, and customize Web Part Pages in document libraries.
     */
    PermissionKind[PermissionKind["EditListItems"] = 3] = "EditListItems";
    /**
     * Delete items from a list, documents from a document library, and Web discussion
     * comments in documents.
     */
    PermissionKind[PermissionKind["DeleteListItems"] = 4] = "DeleteListItems";
    /**
     * Approve a minor version of a list item or document.
     */
    PermissionKind[PermissionKind["ApproveItems"] = 5] = "ApproveItems";
    /**
     * View the source of documents with server-side file handlers.
     */
    PermissionKind[PermissionKind["OpenItems"] = 6] = "OpenItems";
    /**
     * View past versions of a list item or document.
     */
    PermissionKind[PermissionKind["ViewVersions"] = 7] = "ViewVersions";
    /**
     * Delete past versions of a list item or document.
     */
    PermissionKind[PermissionKind["DeleteVersions"] = 8] = "DeleteVersions";
    /**
     * Discard or check in a document which is checked out to another user.
     */
    PermissionKind[PermissionKind["CancelCheckout"] = 9] = "CancelCheckout";
    /**
     * Create, change, and delete personal views of lists.
     */
    PermissionKind[PermissionKind["ManagePersonalViews"] = 10] = "ManagePersonalViews";
    /**
     * Create and delete lists, add or remove columns in a list, and add or remove public views of a list.
     */
    PermissionKind[PermissionKind["ManageLists"] = 12] = "ManageLists";
    /**
     * View forms, views, and application pages, and enumerate lists.
     */
    PermissionKind[PermissionKind["ViewFormPages"] = 13] = "ViewFormPages";
    /**
     * Make content of a list or document library retrieveable for anonymous users through SharePoint search.
     * The list permissions in the site do not change.
     */
    PermissionKind[PermissionKind["AnonymousSearchAccessList"] = 14] = "AnonymousSearchAccessList";
    /**
     * Allow users to open a Site, list, or folder to access items inside that container.
     */
    PermissionKind[PermissionKind["Open"] = 17] = "Open";
    /**
     * View pages in a Site.
     */
    PermissionKind[PermissionKind["ViewPages"] = 18] = "ViewPages";
    /**
     * Add, change, or delete HTML pages or Web Part Pages, and edit the Site using
     * a Windows SharePoint Services compatible editor.
     */
    PermissionKind[PermissionKind["AddAndCustomizePages"] = 19] = "AddAndCustomizePages";
    /**
     * Apply a theme or borders to the entire Site.
     */
    PermissionKind[PermissionKind["ApplyThemeAndBorder"] = 20] = "ApplyThemeAndBorder";
    /**
     * Apply a style sheet (.css file) to the Site.
     */
    PermissionKind[PermissionKind["ApplyStyleSheets"] = 21] = "ApplyStyleSheets";
    /**
     * View reports on Site usage.
     */
    PermissionKind[PermissionKind["ViewUsageData"] = 22] = "ViewUsageData";
    /**
     * Create a Site using Self-Service Site Creation.
     */
    PermissionKind[PermissionKind["CreateSSCSite"] = 23] = "CreateSSCSite";
    /**
     * Create subsites such as team sites, Meeting Workspace sites, and Document Workspace sites.
     */
    PermissionKind[PermissionKind["ManageSubwebs"] = 24] = "ManageSubwebs";
    /**
     * Create a group of users that can be used anywhere within the site collection.
     */
    PermissionKind[PermissionKind["CreateGroups"] = 25] = "CreateGroups";
    /**
     * Create and change permission levels on the Site and assign permissions to users
     * and groups.
     */
    PermissionKind[PermissionKind["ManagePermissions"] = 26] = "ManagePermissions";
    /**
     * Enumerate files and folders in a Site using Microsoft Office SharePoint Designer
     * and WebDAV interfaces.
     */
    PermissionKind[PermissionKind["BrowseDirectories"] = 27] = "BrowseDirectories";
    /**
     * View information about users of the Site.
     */
    PermissionKind[PermissionKind["BrowseUserInfo"] = 28] = "BrowseUserInfo";
    /**
     * Add or remove personal Web Parts on a Web Part Page.
     */
    PermissionKind[PermissionKind["AddDelPrivateWebParts"] = 29] = "AddDelPrivateWebParts";
    /**
     * Update Web Parts to display personalized information.
     */
    PermissionKind[PermissionKind["UpdatePersonalWebParts"] = 30] = "UpdatePersonalWebParts";
    /**
     * Grant the ability to perform all administration tasks for the Site as well as
     * manage content, activate, deactivate, or edit properties of Site scoped Features
     * through the object model or through the user interface (UI). When granted on the
     * root Site of a Site Collection, activate, deactivate, or edit properties of
     * site collection scoped Features through the object model. To browse to the Site
     * Collection Features page and activate or deactivate Site Collection scoped Features
     * through the UI, you must be a Site Collection administrator.
     */
    PermissionKind[PermissionKind["ManageWeb"] = 31] = "ManageWeb";
    /**
     * Content of lists and document libraries in the Web site will be retrieveable for anonymous users through
     * SharePoint search if the list or document library has AnonymousSearchAccessList set.
     */
    PermissionKind[PermissionKind["AnonymousSearchAccessWebLists"] = 32] = "AnonymousSearchAccessWebLists";
    /**
     * Use features that launch client applications. Otherwise, users must work on documents
     * locally and upload changes.
     */
    PermissionKind[PermissionKind["UseClientIntegration"] = 37] = "UseClientIntegration";
    /**
     * Use SOAP, WebDAV, or Microsoft Office SharePoint Designer interfaces to access the Site.
     */
    PermissionKind[PermissionKind["UseRemoteAPIs"] = 38] = "UseRemoteAPIs";
    /**
     * Manage alerts for all users of the Site.
     */
    PermissionKind[PermissionKind["ManageAlerts"] = 39] = "ManageAlerts";
    /**
     * Create e-mail alerts.
     */
    PermissionKind[PermissionKind["CreateAlerts"] = 40] = "CreateAlerts";
    /**
     * Allows a user to change his or her user information, such as adding a picture.
     */
    PermissionKind[PermissionKind["EditMyUserInfo"] = 41] = "EditMyUserInfo";
    /**
     * Enumerate permissions on Site, list, folder, document, or list item.
     */
    PermissionKind[PermissionKind["EnumeratePermissions"] = 63] = "EnumeratePermissions";
    /**
     * Has all permissions on the Site. Not available through the user interface.
     */
    PermissionKind[PermissionKind["FullMask"] = 65] = "FullMask";
})(PermissionKind || (PermissionKind = {}));


/***/ }),

/***/ 3844:
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/security/web.js ***!
  \**********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 9713);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _funcs_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./funcs.js */ 3974);





(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web, "roleDefinitions", _types_js__WEBPACK_IMPORTED_MODULE_2__.RoleDefinitions);
(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web, "roleAssignments", _types_js__WEBPACK_IMPORTED_MODULE_2__.RoleAssignments);
(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web, "firstUniqueAncestorSecurableObject", _spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.SPInstance);
_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web.prototype.getUserEffectivePermissions = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.getUserEffectivePermissions;
_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web.prototype.getCurrentUserEffectivePermissions = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.getCurrentUserEffectivePermissions;
_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web.prototype.breakRoleInheritance = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.breakRoleInheritance;
_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web.prototype.resetRoleInheritance = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.resetRoleInheritance;
_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web.prototype.userHasPermissions = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.userHasPermissions;
_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web.prototype.currentUserHasPermissions = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.currentUserHasPermissions;
_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web.prototype.hasPermissions = _funcs_js__WEBPACK_IMPORTED_MODULE_4__.hasPermissions;


/***/ }),

/***/ 832:
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/sharing/file.js ***!
  \**********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _files_types_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../files/types.js */ 242);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../types.js */ 8713);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 2919);



_files_types_js__WEBPACK_IMPORTED_MODULE_0__._File.prototype.shareWith = async function (loginNames, role = _types_js__WEBPACK_IMPORTED_MODULE_2__.SharingRole.View, requireSignin = false, emailData) {
    const item = await this.getItem();
    return item.shareWith(loginNames, role, requireSignin, emailData);
};
_files_types_js__WEBPACK_IMPORTED_MODULE_0__._File.prototype.getShareLink = async function (kind, expiration = null) {
    const item = await this.getItem();
    return item.getShareLink(kind, expiration);
};
_files_types_js__WEBPACK_IMPORTED_MODULE_0__._File.prototype.checkSharingPermissions = async function (recipients) {
    const item = await this.getItem();
    return item.checkSharingPermissions(recipients);
};
// TODO:: clean up this method signature for next major release
// eslint-disable-next-line max-len
_files_types_js__WEBPACK_IMPORTED_MODULE_0__._File.prototype.getSharingInformation = async function (request = null, expands, selects) {
    const item = await this.getItem();
    return item.getSharingInformation(request, expands, selects);
};
_files_types_js__WEBPACK_IMPORTED_MODULE_0__._File.prototype.getObjectSharingSettings = async function (useSimplifiedRoles = true) {
    const item = await this.getItem();
    return item.getObjectSharingSettings(useSimplifiedRoles);
};
_files_types_js__WEBPACK_IMPORTED_MODULE_0__._File.prototype.unshare = async function () {
    const item = await this.getItem();
    return item.unshare();
};
_files_types_js__WEBPACK_IMPORTED_MODULE_0__._File.prototype.deleteSharingLinkByKind = async function (linkKind) {
    const item = await this.getItem();
    return item.deleteSharingLinkByKind(linkKind);
};
_files_types_js__WEBPACK_IMPORTED_MODULE_0__._File.prototype.unshareLink = async function unshareLink(linkKind, shareId = _types_js__WEBPACK_IMPORTED_MODULE_1__.emptyGuid) {
    const item = await this.getItem();
    return item.unshareLink(linkKind, shareId);
};


/***/ }),

/***/ 2284:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/sharing/folder.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _folders_types_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../folders/types.js */ 187);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 2919);


const field = "odata.id";
_folders_types_js__WEBPACK_IMPORTED_MODULE_0__._Folder.prototype.shareWith = async function (loginNames, role = _types_js__WEBPACK_IMPORTED_MODULE_1__.SharingRole.View, requireSignin = false, emailData) {
    const shareable = await this.getItem(field);
    return shareable.shareWith(loginNames, role, requireSignin, emailData);
};
_folders_types_js__WEBPACK_IMPORTED_MODULE_0__._Folder.prototype.getShareLink = async function (kind, expiration = null) {
    const shareable = await this.getItem(field);
    return shareable.getShareLink(kind, expiration);
};
_folders_types_js__WEBPACK_IMPORTED_MODULE_0__._Folder.prototype.checkSharingPermissions = async function (recipients) {
    const shareable = await this.getItem(field);
    return shareable.checkSharingPermissions(recipients);
};
_folders_types_js__WEBPACK_IMPORTED_MODULE_0__._Folder.prototype.getSharingInformation = async function (request, expands, selects) {
    const shareable = await this.getItem(field);
    return shareable.getSharingInformation(request, expands, selects);
};
_folders_types_js__WEBPACK_IMPORTED_MODULE_0__._Folder.prototype.getObjectSharingSettings = async function (useSimplifiedRoles = true) {
    const shareable = await this.getItem(field);
    return shareable.getObjectSharingSettings(useSimplifiedRoles);
};
_folders_types_js__WEBPACK_IMPORTED_MODULE_0__._Folder.prototype.unshare = async function () {
    const shareable = await this.getItem(field);
    return shareable.unshare();
};
_folders_types_js__WEBPACK_IMPORTED_MODULE_0__._Folder.prototype.deleteSharingLinkByKind = async function (kind) {
    const shareable = await this.getItem(field);
    return shareable.deleteSharingLinkByKind(kind);
};
_folders_types_js__WEBPACK_IMPORTED_MODULE_0__._Folder.prototype.unshareLink = async function (kind, shareId) {
    const shareable = await this.getItem(field);
    return shareable.unshareLink(kind, shareId);
};


/***/ }),

/***/ 2251:
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/sharing/funcs.js ***!
  \***********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   checkPermissions: () => (/* binding */ checkPermissions),
/* harmony export */   deleteLinkByKind: () => (/* binding */ deleteLinkByKind),
/* harmony export */   getObjectSharingSettings: () => (/* binding */ getObjectSharingSettings),
/* harmony export */   getShareLink: () => (/* binding */ getShareLink),
/* harmony export */   getSharingInformation: () => (/* binding */ getSharingInformation),
/* harmony export */   shareObject: () => (/* binding */ shareObject),
/* harmony export */   shareWith: () => (/* binding */ shareWith),
/* harmony export */   unshareLink: () => (/* binding */ unshareLink),
/* harmony export */   unshareObject: () => (/* binding */ unshareObject)
/* harmony export */ });
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../utils/extract-web-url.js */ 2997);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./types.js */ 2919);
/* harmony import */ var _security_types_js__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../security/types.js */ 9713);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../types.js */ 8713);








/**
 * Shares an object based on the supplied options
 *
 * @param options The set of options to send to the ShareObject method
 * @param bypass If true any processing is skipped and the options are sent directly to the ShareObject method
 */
async function shareObject(o, options, bypass = false) {
    if (bypass) {
        // if the bypass flag is set send the supplied parameters directly to the service
        return sendShareObjectRequest(o, options);
    }
    // extend our options with some defaults
    options = {
        group: null,
        includeAnonymousLinkInEmail: false,
        propagateAcl: false,
        useSimplifiedRoles: true,
        ...options,
    };
    const roleValue = await getRoleValue.apply(o, [options.role, options.group]);
    // handle the multiple input types
    if (!(0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.isArray)(options.loginNames)) {
        options.loginNames = [options.loginNames];
    }
    const userStr = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.jsS)(options.loginNames.map(Key => ({ Key })));
    let postBody = {
        peoplePickerInput: userStr,
        roleValue: roleValue,
        url: options.url,
    };
    if (options.emailData !== undefined && options.emailData !== null) {
        postBody = {
            emailBody: options.emailData.body,
            emailSubject: options.emailData.subject !== undefined ? options.emailData.subject : "Shared with you.",
            sendEmail: true,
            ...postBody,
        };
    }
    return sendShareObjectRequest(o, postBody);
}
/**
 * Gets a sharing link for the supplied
 *
 * @param kind The kind of link to share
 * @param expiration The optional expiration for this link
 */
function getShareLink(kind, expiration = null) {
    // date needs to be an ISO string or null
    const expString = expiration !== null ? expiration.toISOString() : null;
    // clone using the factory and send the request
    const o = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPInstance)(this, "shareLink");
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(o, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({
        request: {
            createLink: true,
            emailData: null,
            settings: {
                expiration: expString,
                linkKind: kind,
            },
        },
    }));
}
/**
 * Checks Permissions on the list of Users and returns back role the users have on the Item.
 *
 * @param recipients The array of Entities for which Permissions need to be checked.
 */
function checkPermissions(recipients) {
    const o = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPInstance)(this, "checkPermissions");
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(o, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({ recipients }));
}
/**
 * Get Sharing Information.
 *
 * @param request The SharingInformationRequest Object.
 * @param expands Expand more fields.
 *
 */
function getSharingInformation(request = null, expands = [], selects = ["*"]) {
    const o = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPInstance)(this, "getSharingInformation");
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(o.select(...selects).expand(...expands), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({ request }));
}
/**
 * Gets the sharing settings of an item.
 *
 * @param useSimplifiedRoles Determines whether to use simplified roles.
 */
function getObjectSharingSettings(useSimplifiedRoles = true) {
    const o = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPInstance)(this, "getObjectSharingSettings");
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(o, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({ useSimplifiedRoles }));
}
/**
 * Unshares this object
 */
function unshareObject() {
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)((0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPInstance)(this, "unshareObject"));
}
/**
 * Deletes a link by type
 *
 * @param kind Deletes a sharing link by the kind of link
 */
function deleteLinkByKind(linkKind) {
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)((0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPInstance)(this, "deleteLinkByKind"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({ linkKind }));
}
/**
 * Removes the specified link to the item.
 *
 * @param kind The kind of link to be deleted.
 * @param shareId
 */
function unshareLink(linkKind, shareId = _types_js__WEBPACK_IMPORTED_MODULE_7__.emptyGuid) {
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)((0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPInstance)(this, "unshareLink"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({ linkKind, shareId }));
}
/**
 * Shares this instance with the supplied users
 *
 * @param loginNames Resolved login names to share
 * @param role The role
 * @param requireSignin True to require the user is authenticated, otherwise false
 * @param propagateAcl True to apply this share to all children
 * @param emailData If supplied an email will be sent with the indicated properties
 */
async function shareWith(o, loginNames, role, requireSignin = false, propagateAcl = false, emailData) {
    // handle the multiple input types
    if (!(0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.isArray)(loginNames)) {
        loginNames = [loginNames];
    }
    const userStr = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.jsS)(loginNames.map(login => {
        return { Key: login };
    }));
    const roleFilter = role === _types_js__WEBPACK_IMPORTED_MODULE_5__.SharingRole.Edit ? _types_js__WEBPACK_IMPORTED_MODULE_5__.RoleType.Contributor : _types_js__WEBPACK_IMPORTED_MODULE_5__.RoleType.Reader;
    // start by looking up the role definition id we need to set the roleValue
    const def = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPCollection)([o, (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_3__.extractWebUrl)(o.toUrl())], "_api/web/roledefinitions").select("Id").filter(`RoleTypeKind eq ${roleFilter}`)();
    if (!(0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.isArray)(def) || def.length < 1) {
        throw Error(`Could not locate a role defintion with RoleTypeKind ${roleFilter}`);
    }
    let postBody = {
        includeAnonymousLinkInEmail: requireSignin,
        peoplePickerInput: userStr,
        propagateAcl: propagateAcl,
        roleValue: `role:${def[0].Id}`,
        useSimplifiedRoles: true,
    };
    if (emailData !== undefined) {
        postBody = {
            ...postBody,
            emailBody: emailData.body,
            emailSubject: emailData.subject !== undefined ? emailData.subject : "",
            sendEmail: true,
        };
    }
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)((0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.SPInstance)(o, "shareObject"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)(postBody));
}
async function sendShareObjectRequest(o, options) {
    const w = (0,_webs_types_js__WEBPACK_IMPORTED_MODULE_4__.Web)([o, (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_3__.extractWebUrl)(o.toUrl())], "/_api/SP.Web.ShareObject");
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_2__.spPost)(w.expand("UsersWithAccessRequests", "GroupsSharedWith"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)(options));
}
/**
 * Calculates the roleValue string used in the sharing query
 *
 * @param role The Sharing Role
 * @param group The Group type
 */
async function getRoleValue(role, group) {
    // we will give group precedence, because we had to make a choice
    if (group !== undefined && group !== null) {
        switch (group) {
            case _types_js__WEBPACK_IMPORTED_MODULE_5__.RoleType.Contributor: {
                const g1 = await (0,_webs_types_js__WEBPACK_IMPORTED_MODULE_4__.Web)([this, "_api/web"], "associatedmembergroup").select("Id")();
                return `group: ${g1.Id}`;
            }
            case _types_js__WEBPACK_IMPORTED_MODULE_5__.RoleType.Reader:
            case _types_js__WEBPACK_IMPORTED_MODULE_5__.RoleType.Guest: {
                const g2 = await (0,_webs_types_js__WEBPACK_IMPORTED_MODULE_4__.Web)([this, "_api/web"], "associatedvisitorgroup").select("Id")();
                return `group: ${g2.Id}`;
            }
            default:
                throw Error("Could not determine role value for supplied value. Contributor, Reader, and Guest are supported");
        }
    }
    else {
        const roleFilter = role === _types_js__WEBPACK_IMPORTED_MODULE_5__.SharingRole.Edit ? _types_js__WEBPACK_IMPORTED_MODULE_5__.RoleType.Contributor : _types_js__WEBPACK_IMPORTED_MODULE_5__.RoleType.Reader;
        const def = await (0,_security_types_js__WEBPACK_IMPORTED_MODULE_6__.RoleDefinitions)([this, "_api/web"]).select("Id").top(1).filter(`RoleTypeKind eq ${roleFilter}`)();
        if (def === undefined || (def === null || def === void 0 ? void 0 : def.length) < 1) {
            throw Error("Could not locate associated role definition for supplied role. Edit and View are supported");
        }
        return `role: ${def[0].Id}`;
    }
}


/***/ }),

/***/ 8672:
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/sharing/index.js ***!
  \***********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _file_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./file.js */ 832);
/* harmony import */ var _folder_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./folder.js */ 2284);
/* harmony import */ var _item_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./item.js */ 1557);
/* harmony import */ var _web_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./web.js */ 778);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./types.js */ 2919);







/***/ }),

/***/ 1557:
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/sharing/item.js ***!
  \**********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _items_types_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../items/types.js */ 132);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 2919);
/* harmony import */ var _funcs_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./funcs.js */ 2251);



_items_types_js__WEBPACK_IMPORTED_MODULE_0__._Item.prototype.shareWith = function (loginNames, role = _types_js__WEBPACK_IMPORTED_MODULE_1__.SharingRole.View, requireSignin = false, emailData) {
    return (0,_funcs_js__WEBPACK_IMPORTED_MODULE_2__.shareWith)(this, loginNames, role, requireSignin, false, emailData);
};
_items_types_js__WEBPACK_IMPORTED_MODULE_0__._Item.prototype.getShareLink = _funcs_js__WEBPACK_IMPORTED_MODULE_2__.getShareLink;
_items_types_js__WEBPACK_IMPORTED_MODULE_0__._Item.prototype.checkSharingPermissions = _funcs_js__WEBPACK_IMPORTED_MODULE_2__.checkPermissions;
_items_types_js__WEBPACK_IMPORTED_MODULE_0__._Item.prototype.getSharingInformation = _funcs_js__WEBPACK_IMPORTED_MODULE_2__.getSharingInformation;
_items_types_js__WEBPACK_IMPORTED_MODULE_0__._Item.prototype.getObjectSharingSettings = _funcs_js__WEBPACK_IMPORTED_MODULE_2__.getObjectSharingSettings;
_items_types_js__WEBPACK_IMPORTED_MODULE_0__._Item.prototype.unshare = _funcs_js__WEBPACK_IMPORTED_MODULE_2__.unshareObject;
_items_types_js__WEBPACK_IMPORTED_MODULE_0__._Item.prototype.deleteSharingLinkByKind = _funcs_js__WEBPACK_IMPORTED_MODULE_2__.deleteLinkByKind;
_items_types_js__WEBPACK_IMPORTED_MODULE_0__._Item.prototype.unshareLink = _funcs_js__WEBPACK_IMPORTED_MODULE_2__.unshareLink;


/***/ }),

/***/ 2919:
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/sharing/types.js ***!
  \***********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   RoleType: () => (/* binding */ RoleType),
/* harmony export */   SharingRole: () => (/* binding */ SharingRole)
/* harmony export */ });
/* unused harmony exports SPSharedObjectType, SharingDomainRestrictionMode, SharingOperationStatusCode, SharingLinkKind */
/**
 * Indicates the role of the sharing link
 */
var SharingRole;
(function (SharingRole) {
    SharingRole[SharingRole["None"] = 0] = "None";
    SharingRole[SharingRole["View"] = 1] = "View";
    SharingRole[SharingRole["Edit"] = 2] = "Edit";
    SharingRole[SharingRole["Owner"] = 3] = "Owner";
})(SharingRole || (SharingRole = {}));
var SPSharedObjectType;
(function (SPSharedObjectType) {
    SPSharedObjectType[SPSharedObjectType["Unknown"] = 0] = "Unknown";
    SPSharedObjectType[SPSharedObjectType["File"] = 1] = "File";
    SPSharedObjectType[SPSharedObjectType["Folder"] = 2] = "Folder";
    SPSharedObjectType[SPSharedObjectType["Item"] = 3] = "Item";
    SPSharedObjectType[SPSharedObjectType["List"] = 4] = "List";
    SPSharedObjectType[SPSharedObjectType["Web"] = 5] = "Web";
    SPSharedObjectType[SPSharedObjectType["Max"] = 6] = "Max";
})(SPSharedObjectType || (SPSharedObjectType = {}));
var SharingDomainRestrictionMode;
(function (SharingDomainRestrictionMode) {
    SharingDomainRestrictionMode[SharingDomainRestrictionMode["None"] = 0] = "None";
    SharingDomainRestrictionMode[SharingDomainRestrictionMode["AllowList"] = 1] = "AllowList";
    SharingDomainRestrictionMode[SharingDomainRestrictionMode["BlockList"] = 2] = "BlockList";
})(SharingDomainRestrictionMode || (SharingDomainRestrictionMode = {}));
var SharingOperationStatusCode;
(function (SharingOperationStatusCode) {
    /**
     * The share operation completed without errors.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["CompletedSuccessfully"] = 0] = "CompletedSuccessfully";
    /**
     * The share operation completed and generated requests for access.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["AccessRequestsQueued"] = 1] = "AccessRequestsQueued";
    /**
     * The share operation failed as there were no resolved users.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["NoResolvedUsers"] = -1] = "NoResolvedUsers";
    /**
     * The share operation failed due to insufficient permissions.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["AccessDenied"] = -2] = "AccessDenied";
    /**
     * The share operation failed when attempting a cross site share, which is not supported.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["CrossSiteRequestNotSupported"] = -3] = "CrossSiteRequestNotSupported";
    /**
     * The sharing operation failed due to an unknown error.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["UnknowError"] = -4] = "UnknowError";
    /**
     * The text you typed is too long. Please shorten it.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["EmailBodyTooLong"] = -5] = "EmailBodyTooLong";
    /**
     * The maximum number of unique scopes in the list has been exceeded.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["ListUniqueScopesExceeded"] = -6] = "ListUniqueScopesExceeded";
    /**
     * The share operation failed because a sharing capability is disabled in the site.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["CapabilityDisabled"] = -7] = "CapabilityDisabled";
    /**
     * The specified object for the share operation is not supported.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["ObjectNotSupported"] = -8] = "ObjectNotSupported";
    /**
     * A SharePoint group cannot contain another SharePoint group.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["NestedGroupsNotSupported"] = -9] = "NestedGroupsNotSupported";
})(SharingOperationStatusCode || (SharingOperationStatusCode = {}));
var SharingLinkKind;
(function (SharingLinkKind) {
    /**
     * Uninitialized link
     */
    SharingLinkKind[SharingLinkKind["Uninitialized"] = 0] = "Uninitialized";
    /**
     * Direct link to the object being shared
     */
    SharingLinkKind[SharingLinkKind["Direct"] = 1] = "Direct";
    /**
     * Organization-shareable link to the object being shared with view permissions
     */
    SharingLinkKind[SharingLinkKind["OrganizationView"] = 2] = "OrganizationView";
    /**
     * Organization-shareable link to the object being shared with edit permissions
     */
    SharingLinkKind[SharingLinkKind["OrganizationEdit"] = 3] = "OrganizationEdit";
    /**
     * View only anonymous link
     */
    SharingLinkKind[SharingLinkKind["AnonymousView"] = 4] = "AnonymousView";
    /**
     * Read/Write anonymous link
     */
    SharingLinkKind[SharingLinkKind["AnonymousEdit"] = 5] = "AnonymousEdit";
    /**
     * Flexible sharing Link where properties can change without affecting link URL
     */
    SharingLinkKind[SharingLinkKind["Flexible"] = 6] = "Flexible";
})(SharingLinkKind || (SharingLinkKind = {}));
var RoleType;
(function (RoleType) {
    RoleType[RoleType["None"] = 0] = "None";
    RoleType[RoleType["Guest"] = 1] = "Guest";
    RoleType[RoleType["Reader"] = 2] = "Reader";
    RoleType[RoleType["Contributor"] = 3] = "Contributor";
    RoleType[RoleType["WebDesigner"] = 4] = "WebDesigner";
    RoleType[RoleType["Administrator"] = 5] = "Administrator";
})(RoleType || (RoleType = {}));


/***/ }),

/***/ 778:
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/sharing/web.js ***!
  \*********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 2919);
/* harmony import */ var _funcs_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./funcs.js */ 2251);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../spqueryable.js */ 2678);






/**
 * Shares this web with the supplied users (not supported for batching)
 * @param loginNames The resolved login names to share
 * @param role The role to share this web
 * @param emailData Optional email data
 */
_webs_types_js__WEBPACK_IMPORTED_MODULE_0__._Web.prototype.shareWith = async function (loginNames, role = _types_js__WEBPACK_IMPORTED_MODULE_1__.SharingRole.View, emailData) {
    const url = await this.select("Url")();
    return this.shareObject((0,_pnp_core__WEBPACK_IMPORTED_MODULE_3__.combine)(url.Url, "/_layouts/15/aclinv.aspx?forSharing=1&mbypass=1"), loginNames, role, emailData);
};
/**
 * Provides direct access to the static web.ShareObject method
 *
 * @param url The url to share
 * @param loginNames Resolved loginnames string[] of a single login name string
 * @param roleValue Role value
 * @param emailData Optional email data
 * @param groupId Optional group id
 * @param propagateAcl
 * @param includeAnonymousLinkInEmail
 * @param useSimplifiedRoles
 */
_webs_types_js__WEBPACK_IMPORTED_MODULE_0__._Web.prototype.shareObject = function (url, loginNames, role, emailData, group, propagateAcl = false, includeAnonymousLinkInEmail = false, useSimplifiedRoles = true) {
    return (0,_funcs_js__WEBPACK_IMPORTED_MODULE_2__.shareObject)(this, {
        emailData: emailData,
        group: group,
        includeAnonymousLinkInEmail: includeAnonymousLinkInEmail,
        loginNames: loginNames,
        propagateAcl: propagateAcl,
        role: role,
        url: url,
        useSimplifiedRoles: useSimplifiedRoles,
    });
};
/**
 * Supplies a method to pass any set of arguments to ShareObject
 *
 * @param options The set of options to send to ShareObject
 */
_webs_types_js__WEBPACK_IMPORTED_MODULE_0__._Web.prototype.shareObjectRaw = function (options) {
    return (0,_funcs_js__WEBPACK_IMPORTED_MODULE_2__.shareObject)(this, options, true);
};
/**
 * Supplies a method to pass any set of arguments to ShareObject
 *
 * @param options The set of options to send to ShareObject
 */
_webs_types_js__WEBPACK_IMPORTED_MODULE_0__._Web.prototype.unshareObject = function (url) {
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_5__.spPost)((0,_webs_types_js__WEBPACK_IMPORTED_MODULE_0__.Web)(this, "unshareObject"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_4__.body)({ url }));
};


/***/ }),

/***/ 4678:
/*!****************************************************!*\
  !*** ./node_modules/@pnp/sp/site-designs/index.js ***!
  \****************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _web_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web.js */ 8975);
/* harmony import */ var _fi_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../fi.js */ 6553);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 9288);




Reflect.defineProperty(_fi_js__WEBPACK_IMPORTED_MODULE_1__.SPFI.prototype, "siteDesigns", {
    configurable: true,
    enumerable: true,
    get: function () {
        return this.create(_types_js__WEBPACK_IMPORTED_MODULE_2__.SiteDesigns);
    },
});


/***/ }),

/***/ 9288:
/*!****************************************************!*\
  !*** ./node_modules/@pnp/sp/site-designs/types.js ***!
  \****************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   SiteDesigns: () => (/* binding */ SiteDesigns)
/* harmony export */ });
/* unused harmony exports _SiteDesigns, TemplateDesignType, ListDesignColor, ListDesignIcon */
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../utils/extract-web-url.js */ 2997);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/core */ 1971);




class _SiteDesigns extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPQueryable {
    constructor(base, methodName = "") {
        super(base);
        this._url = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_3__.combine)((0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_1__.extractWebUrl)(this._url), `_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.${methodName}`);
    }
    run(props) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_2__.body)(props, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_2__.headers)({ "Content-Type": "application/json;charset=utf-8" })));
    }
    /**
     * Creates a new site design available to users when they create a new site from the SharePoint home page.
     *
     * @param creationInfo A sitedesign creation information object
     */
    createSiteDesign(creationInfo) {
        return SiteDesignsCloneFactory(this, "CreateSiteDesign").run({ info: creationInfo });
    }
    /**
     * Applies a site design to an existing site collection.
     *
     * @param siteDesignId The ID of the site design to apply.
     * @param webUrl The URL of the site collection where you want to apply the site design.
     */
    applySiteDesign(siteDesignId, webUrl) {
        return SiteDesignsCloneFactory(this, "ApplySiteDesign").run({ siteDesignId: siteDesignId, "webUrl": webUrl });
    }
    /**
     * Gets the list of available site designs
     */
    getSiteDesigns() {
        return SiteDesignsCloneFactory(this, "GetSiteDesigns").run({});
    }
    /**
     * Gets information about a specific site design.
     * @param id The ID of the site design to get information about.
     */
    getSiteDesignMetadata(id) {
        return SiteDesignsCloneFactory(this, "GetSiteDesignMetadata").run({ id: id });
    }
    /**
     * Updates a site design with new values. In the REST call, all parameters are optional except the site script Id.
     * If you had previously set the IsDefault parameter to TRUE and wish it to remain true, you must pass in this parameter again (otherwise it will be reset to FALSE).
     * @param updateInfo A sitedesign update information object
     */
    updateSiteDesign(updateInfo) {
        return SiteDesignsCloneFactory(this, "UpdateSiteDesign").run({ updateInfo: updateInfo });
    }
    /**
     * Deletes a site design.
     * @param id The ID of the site design to delete.
     */
    deleteSiteDesign(id) {
        return SiteDesignsCloneFactory(this, "DeleteSiteDesign").run({ id: id });
    }
    /**
     * Gets a list of principals that have access to a site design.
     * @param id The ID of the site design to get rights information from.
     */
    getSiteDesignRights(id) {
        return SiteDesignsCloneFactory(this, "GetSiteDesignRights").run({ id: id });
    }
    /**
     * Grants access to a site design for one or more principals.
     * @param id The ID of the site design to grant rights on.
     * @param principalNames An array of one or more principals to grant view rights.
     *                       Principals can be users or mail-enabled security groups in the form of "alias" or "alias@<domain name>.com"
     * @param grantedRights Always set to 1. This represents the View right.
     */
    grantSiteDesignRights(id, principalNames, grantedRights = 1) {
        return SiteDesignsCloneFactory(this, "GrantSiteDesignRights").run({
            "grantedRights": grantedRights.toString(),
            id,
            principalNames,
        });
    }
    /**
     * Revokes access from a site design for one or more principals.
     * @param id The ID of the site design to revoke rights from.
     * @param principalNames An array of one or more principals to revoke view rights from.
     *                       If all principals have rights revoked on the site design, the site design becomes viewable to everyone.
     */
    revokeSiteDesignRights(id, principalNames) {
        return SiteDesignsCloneFactory(this, "RevokeSiteDesignRights").run({
            id,
            principalNames,
        });
    }
    /**
     * Adds a site design task on the specified web url to be invoked asynchronously.
     * @param webUrl The absolute url of the web on where to create the task
     * @param siteDesignId The ID of the site design to create a task for
     */
    addSiteDesignTask(webUrl, siteDesignId) {
        return SiteDesignsCloneFactory(this, "AddSiteDesignTask").run({ webUrl, siteDesignId });
    }
    /**
     * Adds a site design task on the current web to be invoked asynchronously.
     * @param siteDesignId The ID of the site design to create a task for
     */
    addSiteDesignTaskToCurrentWeb(siteDesignId) {
        return SiteDesignsCloneFactory(this, "AddSiteDesignTaskToCurrentWeb").run({ siteDesignId });
    }
    /**
     * Retrieves the site design task, if the task has finished running null will be returned
     * @param id The ID of the site design task
     */
    async getSiteDesignTask(id) {
        const task = await SiteDesignsCloneFactory(this, "GetSiteDesignTask").run({ "taskId": id });
        return (0,_pnp_core__WEBPACK_IMPORTED_MODULE_3__.hOP)(task, "ID") ? task : null;
    }
    /**
     * Retrieves a list of site design that have run on a specific web
     * @param webUrl The url of the web where the site design was applied
     * @param siteDesignId (Optional) the site design ID, if not provided will return all site design runs
     */
    getSiteDesignRun(webUrl, siteDesignId) {
        return SiteDesignsCloneFactory(this, "GetSiteDesignRun").run({ webUrl, siteDesignId });
    }
    /**
     * Retrieves the status of a site design that has been run or is still running
     * @param webUrl The url of the web where the site design was applied
     * @param runId the run ID
     */
    getSiteDesignRunStatus(webUrl, runId) {
        return SiteDesignsCloneFactory(this, "GetSiteDesignRunStatus").run({ webUrl, runId });
    }
}
const SiteDesigns = (baseUrl, methodName) => new _SiteDesigns(baseUrl, methodName);
const SiteDesignsCloneFactory = (baseUrl, methodName = "") => SiteDesigns(baseUrl, methodName);
var TemplateDesignType;
(function (TemplateDesignType) {
    /// <summary>
    /// Represents the Site design type.
    /// </summary>
    TemplateDesignType[TemplateDesignType["Site"] = 0] = "Site";
    /// <summary>
    /// Represents the List design type.
    /// </summary>
    TemplateDesignType[TemplateDesignType["List"] = 1] = "List";
})(TemplateDesignType || (TemplateDesignType = {}));
var ListDesignColor;
(function (ListDesignColor) {
    ListDesignColor[ListDesignColor["DarkRed"] = 0] = "DarkRed";
    ListDesignColor[ListDesignColor["Red"] = 1] = "Red";
    ListDesignColor[ListDesignColor["Orange"] = 2] = "Orange";
    ListDesignColor[ListDesignColor["Green"] = 3] = "Green";
    ListDesignColor[ListDesignColor["DarkGreen"] = 4] = "DarkGreen";
    ListDesignColor[ListDesignColor["Teal"] = 5] = "Teal";
    ListDesignColor[ListDesignColor["Blue"] = 6] = "Blue";
    ListDesignColor[ListDesignColor["NavyBlue"] = 7] = "NavyBlue";
    ListDesignColor[ListDesignColor["BluePurple"] = 8] = "BluePurple";
    ListDesignColor[ListDesignColor["DarkBlue"] = 9] = "DarkBlue";
    ListDesignColor[ListDesignColor["Lavendar"] = 10] = "Lavendar";
    ListDesignColor[ListDesignColor["Pink"] = 11] = "Pink";
})(ListDesignColor || (ListDesignColor = {}));
var ListDesignIcon;
(function (ListDesignIcon) {
    ListDesignIcon[ListDesignIcon["Bug"] = 0] = "Bug";
    ListDesignIcon[ListDesignIcon["Calendar"] = 1] = "Calendar";
    ListDesignIcon[ListDesignIcon["BullseyeTarget"] = 2] = "BullseyeTarget";
    ListDesignIcon[ListDesignIcon["ClipboardList"] = 3] = "ClipboardList";
    ListDesignIcon[ListDesignIcon["Airplane"] = 4] = "Airplane";
    ListDesignIcon[ListDesignIcon["Rocket"] = 5] = "Rocket";
    ListDesignIcon[ListDesignIcon["Color"] = 6] = "Color";
    ListDesignIcon[ListDesignIcon["Insights"] = 7] = "Insights";
    ListDesignIcon[ListDesignIcon["CubeShape"] = 8] = "CubeShape";
    ListDesignIcon[ListDesignIcon["TestBeakerSolid"] = 9] = "TestBeakerSolid";
    ListDesignIcon[ListDesignIcon["Robot"] = 10] = "Robot";
    ListDesignIcon[ListDesignIcon["Savings"] = 11] = "Savings";
})(ListDesignIcon || (ListDesignIcon = {}));


/***/ }),

/***/ 8975:
/*!**************************************************!*\
  !*** ./node_modules/@pnp/sp/site-designs/web.js ***!
  \**************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 9288);


_webs_types_js__WEBPACK_IMPORTED_MODULE_0__._Web.prototype.getSiteDesignRuns = function (siteDesignId) {
    return (0,_types_js__WEBPACK_IMPORTED_MODULE_1__.SiteDesigns)(this, "").getSiteDesignRun(undefined, siteDesignId);
};
_webs_types_js__WEBPACK_IMPORTED_MODULE_0__._Web.prototype.addSiteDesignTask = function (siteDesignId) {
    return (0,_types_js__WEBPACK_IMPORTED_MODULE_1__.SiteDesigns)(this, "").addSiteDesignTaskToCurrentWeb(siteDesignId);
};
_webs_types_js__WEBPACK_IMPORTED_MODULE_0__._Web.prototype.getSiteDesignRunStatus = function (runId) {
    return (0,_types_js__WEBPACK_IMPORTED_MODULE_1__.SiteDesigns)(this, "").getSiteDesignRunStatus(undefined, runId);
};


/***/ }),

/***/ 2063:
/*!***************************************************!*\
  !*** ./node_modules/@pnp/sp/site-groups/index.js ***!
  \***************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _web_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web.js */ 9805);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 4281);




/***/ }),

/***/ 4281:
/*!***************************************************!*\
  !*** ./node_modules/@pnp/sp/site-groups/types.js ***!
  \***************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   SiteGroup: () => (/* binding */ SiteGroup),
/* harmony export */   SiteGroups: () => (/* binding */ SiteGroups)
/* harmony export */ });
/* unused harmony exports _SiteGroups, _SiteGroup */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _site_users_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../site-users/types.js */ 2267);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../decorators.js */ 6540);





let _SiteGroups = class _SiteGroups extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPCollection {
    /**
     * Gets a group from the collection by id
     *
     * @param id The id of the group to retrieve
     */
    getById(id) {
        return SiteGroup(this).concat(`(${id})`);
    }
    /**
     * Adds a new group to the site collection
     *
     * @param properties The group properties object of property names and values to be set for the group
     */
    async add(properties) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_2__.body)(properties));
    }
    /**
     * Gets a group from the collection by name
     *
     * @param groupName The name of the group to retrieve
     */
    getByName(groupName) {
        return SiteGroup(this, `getByName('${groupName}')`);
    }
    /**
     * Removes the group with the specified member id from the collection
     *
     * @param id The id of the group to remove
     */
    removeById(id) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(SiteGroups(this, `removeById('${id}')`));
    }
    /**
     * Removes the cross-site group with the specified name from the collection
     *
     * @param loginName The name of the group to remove
     */
    removeByLoginName(loginName) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(SiteGroups(this, `removeByLoginName('${loginName}')`));
    }
};
_SiteGroups = (0,tslib__WEBPACK_IMPORTED_MODULE_3__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_4__.defaultPath)("sitegroups")
], _SiteGroups);

const SiteGroups = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_SiteGroups);
class _SiteGroup extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPInstance {
    /**
     * Gets the users for this group
     *
     */
    get users() {
        return (0,_site_users_types_js__WEBPACK_IMPORTED_MODULE_1__.SiteUsers)(this, "users");
    }
    /**
    * @param props Group properties to update
    */
    async update(props) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPostMerge)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_2__.body)(props));
    }
    /**
     * Set the owner of a group using a user id
     * @param userId the id of the user that will be set as the owner of the current group
     */
    setUserAsOwner(userId) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(SiteGroup(this, `SetUserAsOwner(${userId})`));
    }
}
const SiteGroup = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_SiteGroup);


/***/ }),

/***/ 9805:
/*!*************************************************!*\
  !*** ./node_modules/@pnp/sp/site-groups/web.js ***!
  \*************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./types.js */ 4281);
/* harmony import */ var _security_web_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../security/web.js */ 3844);





(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_2__._Web, "siteGroups", _types_js__WEBPACK_IMPORTED_MODULE_3__.SiteGroups);
(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_2__._Web, "associatedOwnerGroup", _types_js__WEBPACK_IMPORTED_MODULE_3__.SiteGroup);
(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_2__._Web, "associatedMemberGroup", _types_js__WEBPACK_IMPORTED_MODULE_3__.SiteGroup);
(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_2__._Web, "associatedVisitorGroup", _types_js__WEBPACK_IMPORTED_MODULE_3__.SiteGroup);
_webs_types_js__WEBPACK_IMPORTED_MODULE_2__._Web.prototype.createDefaultAssociatedGroups = async function (groupNameSeed, siteOwner, copyRoleAssignments = false, clearSubscopes = true, siteOwner2) {
    await this.breakRoleInheritance(copyRoleAssignments, clearSubscopes);
    const q = (0,_webs_types_js__WEBPACK_IMPORTED_MODULE_2__.Web)(this, "createDefaultAssociatedGroups(userLogin=@u,userLogin2=@v,groupNameSeed=@s)");
    q.query.set("@u", `'${siteOwner || ""}'`);
    q.query.set("@v", `'${siteOwner2 || ""}'`);
    q.query.set("@s", `'${groupNameSeed || ""}'`);
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(q);
};


/***/ }),

/***/ 4792:
/*!****************************************************!*\
  !*** ./node_modules/@pnp/sp/site-scripts/index.js ***!
  \****************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _web_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web.js */ 5596);
/* harmony import */ var _list_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./list.js */ 4612);
/* harmony import */ var _fi_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../fi.js */ 6553);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./types.js */ 2605);





Reflect.defineProperty(_fi_js__WEBPACK_IMPORTED_MODULE_2__.SPFI.prototype, "siteScripts", {
    configurable: true,
    enumerable: true,
    get: function () {
        return this.create(_types_js__WEBPACK_IMPORTED_MODULE_3__.SiteScripts);
    },
});


/***/ }),

/***/ 4612:
/*!***************************************************!*\
  !*** ./node_modules/@pnp/sp/site-scripts/list.js ***!
  \***************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _lists_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types.js */ 9706);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 2605);
/* harmony import */ var _folders_list_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../folders/list.js */ 6488);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../utils/extract-web-url.js */ 2997);






_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List.prototype.getSiteScript = async function () {
    const rootFolder = await (0,_lists_types_js__WEBPACK_IMPORTED_MODULE_1__.List)(this).rootFolder();
    const web = await (0,_webs_types_js__WEBPACK_IMPORTED_MODULE_4__.Web)([this, (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_5__.extractWebUrl)(this.toUrl())]).select("Url")();
    const absoluteListUrl = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.combine)(web.Url, "Lists", rootFolder.Name);
    return (0,_types_js__WEBPACK_IMPORTED_MODULE_2__.SiteScripts)(this, "").getSiteScriptFromList(absoluteListUrl);
};


/***/ }),

/***/ 2605:
/*!****************************************************!*\
  !*** ./node_modules/@pnp/sp/site-scripts/types.js ***!
  \****************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   SiteScripts: () => (/* binding */ SiteScripts)
/* harmony export */ });
/* unused harmony exports _SiteScripts, SiteScriptActionOutcome */
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../utils/extract-web-url.js */ 2997);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../utils/encode-path-str.js */ 4729);





class _SiteScripts extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPQueryable {
    constructor(base, methodName = "") {
        super(base);
        this._url = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_3__.combine)((0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_2__.extractWebUrl)(this._url), `_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.${methodName}`);
    }
    run(props) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)(props));
    }
    /**
     * Gets a list of information on all existing site scripts.
     */
    getSiteScripts() {
        return SiteScriptsCloneFactory(this, "GetSiteScripts").run({});
    }
    /**
     * Creates a new site script.
     *
     * @param title The display name of the site script.
     * @param content JSON value that describes the script. For more information, see JSON reference.
     */
    createSiteScript(title, description, content) {
        return SiteScriptsCloneFactory(this, `CreateSiteScript(Title=@title,Description=@desc)?@title='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_4__.encodePath)(title)}'&@desc='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_4__.encodePath)(description)}'`)
            .run(content);
    }
    /**
     * Gets information about a specific site script. It also returns the JSON of the script.
     *
     * @param id The ID of the site script to get information about.
     */
    getSiteScriptMetadata(id) {
        return SiteScriptsCloneFactory(this, "GetSiteScriptMetadata").run({ id });
    }
    /**
     * Deletes a site script.
     *
     * @param id The ID of the site script to delete.
     */
    deleteSiteScript(id) {
        return SiteScriptsCloneFactory(this, "DeleteSiteScript").run({ id });
    }
    /**
     * Updates a site script with new values. In the REST call, all parameters are optional except the site script Id.
     *
     * @param siteScriptUpdateInfo Object that contains the information to update a site script.
     *                             Make sure you stringify the content object or pass it in the second 'content' parameter
     * @param content (Optional) A new JSON script defining the script actions. For more information, see Site design JSON schema.
     */
    updateSiteScript(updateInfo, content) {
        if (content) {
            updateInfo.Content = JSON.stringify(content);
        }
        return SiteScriptsCloneFactory(this, "UpdateSiteScript").run({ updateInfo });
    }
    /**
     * Gets the site script syntax (JSON) for a specific list
     * @param listUrl The absolute url of the list to retrieve site script
     */
    getSiteScriptFromList(listUrl) {
        return SiteScriptsCloneFactory(this, "GetSiteScriptFromList").run({ listUrl });
    }
    /**
     * Gets the site script syntax (JSON) for a specific web
     * @param webUrl The absolute url of the web to retrieve site script
     * @param extractInfo configuration object to specify what to extract
     */
    getSiteScriptFromWeb(webUrl, info) {
        return SiteScriptsCloneFactory(this, "getSiteScriptFromWeb").run({ webUrl, info });
    }
    /**
     * Executes the indicated site design action on the indicated web.
     *
     * @param webUrl The absolute url of the web to retrieve site script
     * @param extractInfo configuration object to specify what to extract
     */
    executeSiteScriptAction(actionDefinition) {
        return SiteScriptsCloneFactory(this, "executeSiteScriptAction").run({ actionDefinition });
    }
}
const SiteScripts = (baseUrl, methodName) => new _SiteScripts(baseUrl, methodName);
const SiteScriptsCloneFactory = (baseUrl, methodName = "") => SiteScripts(baseUrl, methodName);
var SiteScriptActionOutcome;
(function (SiteScriptActionOutcome) {
    /**
     * The stage was deemed to have completed successfully.
     */
    SiteScriptActionOutcome[SiteScriptActionOutcome["Success"] = 0] = "Success";
    /**
     * The stage was deemed to have failed to complete successfully (non-blocking, rest of recipe
     * execution should still be able to proceed).
     */
    SiteScriptActionOutcome[SiteScriptActionOutcome["Failure"] = 1] = "Failure";
    /**
     * No action was taken for this stage / this stage was skipped.
     */
    SiteScriptActionOutcome[SiteScriptActionOutcome["NoOp"] = 2] = "NoOp";
    /**
     * There was an exception but the operation succeeded. This is analagous to the operation completing
     * in a "yellow" state.
     */
    SiteScriptActionOutcome[SiteScriptActionOutcome["SucceededWithException"] = 3] = "SucceededWithException";
})(SiteScriptActionOutcome || (SiteScriptActionOutcome = {}));


/***/ }),

/***/ 5596:
/*!**************************************************!*\
  !*** ./node_modules/@pnp/sp/site-scripts/web.js ***!
  \**************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 2605);



_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web.prototype.getSiteScript = async function (extractInfo) {
    const info = await this.select("Url")();
    return (0,_types_js__WEBPACK_IMPORTED_MODULE_2__.SiteScripts)(this.toUrl(), "").using((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.AssignFrom)(this)).getSiteScriptFromWeb(info.Url, extractInfo);
};


/***/ }),

/***/ 1157:
/*!**************************************************!*\
  !*** ./node_modules/@pnp/sp/site-users/index.js ***!
  \**************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _web_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web.js */ 6661);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 2267);




/***/ }),

/***/ 2267:
/*!**************************************************!*\
  !*** ./node_modules/@pnp/sp/site-users/types.js ***!
  \**************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   SiteUser: () => (/* binding */ SiteUser),
/* harmony export */   SiteUsers: () => (/* binding */ SiteUsers)
/* harmony export */ });
/* unused harmony exports _SiteUsers, _SiteUser */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _site_groups_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../site-groups/types.js */ 4281);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../decorators.js */ 6540);





let _SiteUsers = class _SiteUsers extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPCollection {
    /**
     * Gets a user from the collection by id
     *
     * @param id The id of the user to retrieve
     */
    getById(id) {
        return SiteUser(this, `getById(${id})`);
    }
    /**
     * Gets a user from the collection by email
     *
     * @param email The email address of the user to retrieve
     */
    getByEmail(email) {
        return SiteUser(this, `getByEmail('${email}')`);
    }
    /**
     * Gets a user from the collection by login name
     *
     * @param loginName The login name of the user to retrieve
     *   e.g. SharePoint Online: 'i:0#.f|membership|user@domain'
     */
    getByLoginName(loginName) {
        return SiteUser(this).concat(`('!@v::${loginName}')`);
    }
    /**
     * Removes a user from the collection by id
     *
     * @param id The id of the user to remove
     */
    removeById(id) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(SiteUsers(this, `removeById(${id})`));
    }
    /**
     * Removes a user from the collection by login name
     *
     * @param loginName The login name of the user to remove
     */
    removeByLoginName(loginName) {
        const o = SiteUsers(this, "removeByLoginName(@v)");
        o.query.set("@v", `'${loginName}'`);
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(o);
    }
    /**
     * Adds a user to a site collection
     *
     * @param loginName The login name of the user to add  to a site collection
     *
     */
    async add(loginName) {
        await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_2__.body)({ LoginName: loginName }));
        return this.getByLoginName(loginName);
    }
};
_SiteUsers = (0,tslib__WEBPACK_IMPORTED_MODULE_3__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_4__.defaultPath)("siteusers")
], _SiteUsers);

const SiteUsers = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_SiteUsers);
/**
 * Describes a single user
 *
 */
class _SiteUser extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPInstance {
    constructor() {
        super(...arguments);
        this.delete = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.deleteable)();
    }
    /**
     * Gets the groups for this user
     *
     */
    get groups() {
        return (0,_site_groups_types_js__WEBPACK_IMPORTED_MODULE_1__.SiteGroups)(this, "groups");
    }
    /**
     * Updates this user
     *
     * @param props Group properties to update
     */
    async update(props) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPostMerge)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_2__.body)(props));
    }
}
const SiteUser = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_SiteUser);


/***/ }),

/***/ 6661:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/site-users/web.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 2267);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../spqueryable.js */ 2678);




(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web, "siteUsers", _types_js__WEBPACK_IMPORTED_MODULE_2__.SiteUsers);
(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web, "currentUser", _types_js__WEBPACK_IMPORTED_MODULE_2__.SiteUser);
_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web.prototype.ensureUser = async function (logonName) {
    return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_3__.spPost)((0,_webs_types_js__WEBPACK_IMPORTED_MODULE_1__.Web)(this, "ensureuser"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({ logonName }));
};
_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web.prototype.getUserById = function (id) {
    return (0,_types_js__WEBPACK_IMPORTED_MODULE_2__.SiteUser)(this, `getUserById(${id})`);
};


/***/ }),

/***/ 5215:
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/sites/index.js ***!
  \*********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _fi_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../fi.js */ 6553);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 1114);



Reflect.defineProperty(_fi_js__WEBPACK_IMPORTED_MODULE_0__.SPFI.prototype, "site", {
    configurable: true,
    enumerable: true,
    get: function () {
        return this.create(_types_js__WEBPACK_IMPORTED_MODULE_1__.Site);
    },
});


/***/ }),

/***/ 1114:
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/sites/types.js ***!
  \*********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Site: () => (/* binding */ Site),
/* harmony export */   _Site: () => (/* binding */ _Site)
/* harmony export */ });
/* unused harmony exports SiteLogoType, SiteLogoAspect */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../decorators.js */ 6540);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../utils/odata-url-from.js */ 905);
/* harmony import */ var _utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../utils/extract-web-url.js */ 2997);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../types.js */ 8713);









/**
 * Ensures that whatever url is passed to the constructor we can correctly rebase it to a site url
 *
 * @param candidate The candidate site url
 * @param path The caller supplied path, which may contain _api, meaning we don't append _api/site
 */
function rebaseSiteUrl(candidate, path) {
    let replace = "_api/site";
    // this allows us to both:
    // - test if `candidate` already has an api path
    // - ensure that we append the correct one as sometimes a web is not defined
    //   by _api/web, in the case of _api/site/rootweb for example
    const matches = /(_api[/|\\](site|web))/i.exec(candidate);
    if ((matches === null || matches === void 0 ? void 0 : matches.length) > 0) {
        // we want just the base url part (before the _api)
        candidate = (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_5__.extractWebUrl)(candidate);
        // we want to ensure we put back the correct string
        replace = matches[1];
    }
    // we only need to append the _api part IF `path` doesn't already include it.
    if ((path === null || path === void 0 ? void 0 : path.indexOf("_api")) < 0) {
        candidate = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_2__.combine)(candidate, replace);
    }
    return candidate;
}
let _Site = class _Site extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPInstance {
    constructor(base, path) {
        if (typeof base === "string") {
            base = rebaseSiteUrl(base, path);
        }
        else if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_2__.isArray)(base)) {
            base = [base[0], rebaseSiteUrl(base[1], path)];
        }
        else {
            base = [base, rebaseSiteUrl(base.toUrl(), path)];
        }
        super(base, path);
    }
    /**
     * Gets the root web of the site collection
     *
     */
    get rootWeb() {
        return (0,_webs_types_js__WEBPACK_IMPORTED_MODULE_1__.Web)(this, "rootweb");
    }
    /**
     * Returns the collection of changes from the change log that have occurred within the list, based on the specified query
     *
     * @param query The change query
     */
    getChanges(query) {
        const postBody = (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_3__.body)({ query });
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)((0,_webs_types_js__WEBPACK_IMPORTED_MODULE_1__.Web)(this, "getchanges"), postBody);
    }
    /**
     * Opens a web by id (using POST)
     *
     * @param webId The GUID id of the web to open
     */
    async openWebById(webId) {
        const data = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(Site(this, `openWebById('${webId}')`));
        return {
            data,
            web: (0,_webs_types_js__WEBPACK_IMPORTED_MODULE_1__.Web)([this, (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_5__.extractWebUrl)((0,_utils_odata_url_from_js__WEBPACK_IMPORTED_MODULE_4__.odataUrlFrom)(data))]),
        };
    }
    /**
     * Gets a Web instance representing the root web of the site collection
     * correctly setup for chaining within the library
     */
    async getRootWeb() {
        const web = await this.rootWeb.select("Url")();
        return (0,_webs_types_js__WEBPACK_IMPORTED_MODULE_1__.Web)([this, web.Url]);
    }
    /**
     * Deletes the current site
     *
     */
    async delete() {
        const site = await Site(this, "").select("Id")();
        const q = Site([this, this.parentUrl], "_api/SPSiteManager/Delete");
        await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(q, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_3__.body)({ siteId: site.Id }));
    }
    /**
     * Gets the document libraries on a site. Static method. (SharePoint Online only)
     *
     * @param absoluteWebUrl The absolute url of the web whose document libraries should be returned
     */
    async getDocumentLibraries(absoluteWebUrl) {
        const q = Site([this, this.parentUrl], "_api/sp.web.getdocumentlibraries(@v)");
        q.query.set("@v", `'${absoluteWebUrl}'`);
        const data = await q();
        return (0,_pnp_core__WEBPACK_IMPORTED_MODULE_2__.hOP)(data, "GetDocumentLibraries") ? data.GetDocumentLibraries : data;
    }
    /**
     * Gets the site url from a page url
     *
     * @param absolutePageUrl The absolute url of the page
     */
    async getWebUrlFromPageUrl(absolutePageUrl) {
        const q = Site([this, this.parentUrl], "_api/sp.web.getweburlfrompageurl(@v)");
        q.query.set("@v", `'${absolutePageUrl}'`);
        const data = await q();
        return (0,_pnp_core__WEBPACK_IMPORTED_MODULE_2__.hOP)(data, "GetWebUrlFromPageUrl") ? data.GetWebUrlFromPageUrl : data;
    }
    /**
     * Creates a Modern communication site.
     *
     * @param title The title of the site to create
     * @param lcid The language to use for the site. If not specified will default to 1033 (English).
     * @param shareByEmailEnabled If set to true, it will enable sharing files via Email. By default it is set to false
     * @param url The fully qualified URL (e.g. https://yourtenant.sharepoint.com/sites/mysitecollection) of the site.
     * @param description The description of the communication site.
     * @param classification The Site classification to use. For instance 'Contoso Classified'. See https://www.youtube.com/watch?v=E-8Z2ggHcS0 for more information
     * @param siteDesignId The Guid of the site design to be used.
     *                     You can use the below default OOTB GUIDs:
     *                     Topic: 00000000-0000-0000-0000-000000000000
     *                     Showcase: 6142d2a0-63a5-4ba0-aede-d9fefca2c767
     *                     Blank: f6cc5403-0d63-442e-96c0-285923709ffc
     * @param hubSiteId The id of the hub site to which the new site should be associated
     * @param owner Optional owner value, required if executing the method in app only mode
     */
    async createCommunicationSite(title, lcid = 1033, shareByEmailEnabled = false, url, description, classification, siteDesignId, hubSiteId, owner) {
        return this.createCommunicationSiteFromProps({
            Classification: classification,
            Description: description,
            HubSiteId: hubSiteId,
            Lcid: lcid,
            Owner: owner,
            ShareByEmailEnabled: shareByEmailEnabled,
            SiteDesignId: siteDesignId,
            Title: title,
            Url: url,
        });
    }
    async createCommunicationSiteFromProps(props) {
        // handle defaults
        const request = {
            Classification: "",
            Description: "",
            HubSiteId: _types_js__WEBPACK_IMPORTED_MODULE_6__.emptyGuid,
            Lcid: 1033,
            ShareByEmailEnabled: false,
            SiteDesignId: _types_js__WEBPACK_IMPORTED_MODULE_6__.emptyGuid,
            WebTemplate: "SITEPAGEPUBLISHING#0",
            WebTemplateExtensionId: _types_js__WEBPACK_IMPORTED_MODULE_6__.emptyGuid,
            ...props,
        };
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(Site([this, (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_5__.extractWebUrl)(this.toUrl())], "/_api/SPSiteManager/Create"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_3__.body)({ request }));
    }
    /**
     *
     * @param url Site Url that you want to check if exists
     */
    async exists(url) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(Site([this, (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_5__.extractWebUrl)(this.toUrl())], "/_api/SP.Site.Exists"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_3__.body)({ url }));
    }
    /**
     * Creates a Modern team site backed by Office 365 group. For use in SP Online only. This will not work with App-only tokens
     *
     * @param displayName The title or display name of the Modern team site to be created
     * @param alias Alias of the underlying Office 365 Group
     * @param isPublic Defines whether the Office 365 Group will be public (default), or private.
     * @param lcid The language to use for the site. If not specified will default to English (1033).
     * @param description The description of the site to be created.
     * @param classification The Site classification to use. For instance 'Contoso Classified'. See https://www.youtube.com/watch?v=E-8Z2ggHcS0 for more information
     * @param owners The Owners of the site to be created
     */
    async createModernTeamSite(displayName, alias, isPublic, lcid, description, classification, owners, hubSiteId, siteDesignId) {
        return this.createModernTeamSiteFromProps({
            alias,
            classification,
            description,
            displayName,
            hubSiteId,
            isPublic,
            lcid,
            owners,
            siteDesignId,
        });
    }
    async createModernTeamSiteFromProps(props) {
        // handle defaults
        const p = Object.assign({}, {
            classification: "",
            description: "",
            hubSiteId: _types_js__WEBPACK_IMPORTED_MODULE_6__.emptyGuid,
            isPublic: true,
            lcid: 1033,
            owners: [],
        }, props);
        const postBody = {
            alias: p.alias,
            displayName: p.displayName,
            isPublic: p.isPublic,
            optionalParams: {
                Classification: p.classification,
                CreationOptions: [`SPSiteLanguage:${p.lcid}`, `HubSiteId:${p.hubSiteId}`],
                Description: p.description,
                Owners: p.owners,
            },
        };
        if (p.siteDesignId) {
            postBody.optionalParams.CreationOptions.push(`implicit_formula_292aa8a00786498a87a5ca52d9f4214a_${p.siteDesignId}`);
        }
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(Site([this, (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_5__.extractWebUrl)(this.toUrl())], "/_api/GroupSiteManager/CreateGroupEx").using((0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_3__.TextParse)()), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_3__.body)(postBody));
    }
    update(props) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPatch)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_3__.body)(props));
    }
    /**
     * Set's the site's `Site Logo` property, vs the Site Icon property available on the web's properties
     *
     * @param logoProperties An instance of ISiteLogoProperties which sets the new site logo.
     */
    setSiteLogo(logoProperties) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)((0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.SPQueryable)([this, (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_5__.extractWebUrl)(this.toUrl())], "_api/siteiconmanager/setsitelogo"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_3__.body)(logoProperties));
    }
};
_Site = (0,tslib__WEBPACK_IMPORTED_MODULE_7__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_8__.defaultPath)("_api/site")
], _Site);

const Site = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_Site);
var SiteLogoType;
(function (SiteLogoType) {
    /**
     * Site header logo
     */
    SiteLogoType[SiteLogoType["WebLogo"] = 0] = "WebLogo";
    /**
     * Hub site logo
     */
    SiteLogoType[SiteLogoType["HubLogo"] = 1] = "HubLogo";
    /**
     * Header background image
     */
    SiteLogoType[SiteLogoType["HeaderBackground"] = 2] = "HeaderBackground";
    /**
     * Global navigation logo
     */
    SiteLogoType[SiteLogoType["GlobalNavLogo"] = 3] = "GlobalNavLogo";
})(SiteLogoType || (SiteLogoType = {}));
var SiteLogoAspect;
(function (SiteLogoAspect) {
    SiteLogoAspect[SiteLogoAspect["Square"] = 0] = "Square";
    SiteLogoAspect[SiteLogoAspect["Rectangular"] = 1] = "Rectangular";
})(SiteLogoAspect || (SiteLogoAspect = {}));


/***/ }),

/***/ 4521:
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/social/index.js ***!
  \**********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./types.js */ 668);
/* harmony import */ var _fi_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../fi.js */ 6553);



Reflect.defineProperty(_fi_js__WEBPACK_IMPORTED_MODULE_1__.SPFI.prototype, "social", {
    configurable: true,
    enumerable: true,
    get: function () {
        return this.create(_types_js__WEBPACK_IMPORTED_MODULE_0__.Social);
    },
});


/***/ }),

/***/ 668:
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/social/types.js ***!
  \**********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Social: () => (/* binding */ Social)
/* harmony export */ });
/* unused harmony exports _Social, _MySocial, MySocial, SocialActorType, SocialActorTypes, SocialFollowResult, SocialStatusCode */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../decorators.js */ 6540);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/queryable */ 6844);





let _Social = class _Social extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPInstance {
    get my() {
        return MySocial(this);
    }
    async getFollowedSitesUri() {
        const r = await SocialCloneFactory(this, "FollowedSitesUri")();
        return r.FollowedSitesUri || r;
    }
    async getFollowedDocumentsUri() {
        const r = await SocialCloneFactory(this, "FollowedDocumentsUri")();
        return r.FollowedDocumentsUri || r;
    }
    async follow(actorInfo) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(SocialCloneFactory(this, "follow"), this.createSocialActorInfoRequestBody(actorInfo));
    }
    async isFollowed(actorInfo) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(SocialCloneFactory(this, "isfollowed"), this.createSocialActorInfoRequestBody(actorInfo));
    }
    async stopFollowing(actorInfo) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(SocialCloneFactory(this, "stopfollowing"), this.createSocialActorInfoRequestBody(actorInfo));
    }
    createSocialActorInfoRequestBody(actorInfo) {
        return (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_2__.body)({
            "actor": {
                Id: null,
                ...actorInfo,
            },
        });
    }
};
_Social = (0,tslib__WEBPACK_IMPORTED_MODULE_3__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_4__.defaultPath)("_api/social.following")
], _Social);

/**
 * Get a new Social instance for the particular Url
 */
const Social = (baseUrl) => new _Social(baseUrl);
const SocialCloneFactory = (baseUrl, paths) => new _Social(baseUrl, paths);
/**
 * Current user's Social instance
 */
let _MySocial = class _MySocial extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPInstance {
    async followed(types) {
        const r = await MySocialCloneFactory(this, `followed(types=${types})`)();
        return (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(r, "Followed") ? r.Followed.results : r;
    }
    async followedCount(types) {
        const r = await MySocialCloneFactory(this, `followedcount(types=${types})`)();
        return r.FollowedCount || r;
    }
    async followers() {
        const r = await MySocialCloneFactory(this, "followers")();
        return (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(r, "Followers") ? r.Followers.results : r;
    }
    async suggestions() {
        const r = await MySocialCloneFactory(this, "suggestions")();
        return (0,_pnp_core__WEBPACK_IMPORTED_MODULE_1__.hOP)(r, "Suggestions") ? r.Suggestions.results : r;
    }
};
_MySocial = (0,tslib__WEBPACK_IMPORTED_MODULE_3__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_4__.defaultPath)("my")
], _MySocial);

/**
 * Invokable factory for IMySocial instances
 */
const MySocial = (baseUrl, path) => new _MySocial(baseUrl, path);
const MySocialCloneFactory = (baseUrl, path) => new _MySocial(baseUrl, path);
/**
 * Social actor type
 *
 */
var SocialActorType;
(function (SocialActorType) {
    SocialActorType[SocialActorType["User"] = 0] = "User";
    SocialActorType[SocialActorType["Document"] = 1] = "Document";
    SocialActorType[SocialActorType["Site"] = 2] = "Site";
    SocialActorType[SocialActorType["Tag"] = 3] = "Tag";
})(SocialActorType || (SocialActorType = {}));
/**
 * Social actor type
 *
 */
/* eslint-disable no-bitwise */
var SocialActorTypes;
(function (SocialActorTypes) {
    SocialActorTypes[SocialActorTypes["None"] = 0] = "None";
    SocialActorTypes[SocialActorTypes["User"] = 1] = "User";
    SocialActorTypes[SocialActorTypes["Document"] = 2] = "Document";
    SocialActorTypes[SocialActorTypes["Site"] = 4] = "Site";
    SocialActorTypes[SocialActorTypes["Tag"] = 8] = "Tag";
    /**
   * The set excludes documents and sites that do not have feeds.
   */
    SocialActorTypes[SocialActorTypes["ExcludeContentWithoutFeeds"] = 268435456] = "ExcludeContentWithoutFeeds";
    /**
   * The set includes group sites
   */
    SocialActorTypes[SocialActorTypes["IncludeGroupsSites"] = 536870912] = "IncludeGroupsSites";
    /**
   * The set includes only items created within the last 24 hours
   */
    SocialActorTypes[SocialActorTypes["WithinLast24Hours"] = 1073741824] = "WithinLast24Hours";
})(SocialActorTypes || (SocialActorTypes = {}));
/* eslint-enable no-bitwise */
/**
 * Result from following
 *
 */
var SocialFollowResult;
(function (SocialFollowResult) {
    SocialFollowResult[SocialFollowResult["Ok"] = 0] = "Ok";
    SocialFollowResult[SocialFollowResult["AlreadyFollowing"] = 1] = "AlreadyFollowing";
    SocialFollowResult[SocialFollowResult["LimitReached"] = 2] = "LimitReached";
    SocialFollowResult[SocialFollowResult["InternalError"] = 3] = "InternalError";
})(SocialFollowResult || (SocialFollowResult = {}));
/**
 * Specifies an exception or status code.
 */
var SocialStatusCode;
(function (SocialStatusCode) {
    /**
   * The operation completed successfully
   */
    SocialStatusCode[SocialStatusCode["OK"] = 0] = "OK";
    /**
   * The request is invalid.
   */
    SocialStatusCode[SocialStatusCode["InvalidRequest"] = 1] = "InvalidRequest";
    /**
   *  The current user is not authorized to perform the operation.
   */
    SocialStatusCode[SocialStatusCode["AccessDenied"] = 2] = "AccessDenied";
    /**
   * The target of the operation was not found.
   */
    SocialStatusCode[SocialStatusCode["ItemNotFound"] = 3] = "ItemNotFound";
    /**
   * The operation is invalid for the target's current state.
   */
    SocialStatusCode[SocialStatusCode["InvalidOperation"] = 4] = "InvalidOperation";
    /**
   * The operation completed without modifying the target.
   */
    SocialStatusCode[SocialStatusCode["ItemNotModified"] = 5] = "ItemNotModified";
    /**
   * The operation failed because an internal error occurred.
   */
    SocialStatusCode[SocialStatusCode["InternalError"] = 6] = "InternalError";
    /**
   * The operation failed because the server could not access the distributed cache.
   */
    SocialStatusCode[SocialStatusCode["CacheReadError"] = 7] = "CacheReadError";
    /**
   * The operation succeeded but the server could not update the distributed cache.
   */
    SocialStatusCode[SocialStatusCode["CacheUpdateError"] = 8] = "CacheUpdateError";
    /**
   * No personal site exists for the current user, and no further information is available.
   */
    SocialStatusCode[SocialStatusCode["PersonalSiteNotFound"] = 9] = "PersonalSiteNotFound";
    /**
   * No personal site exists for the current user, and a previous attempt to create one failed.
   */
    SocialStatusCode[SocialStatusCode["FailedToCreatePersonalSite"] = 10] = "FailedToCreatePersonalSite";
    /**
   * No personal site exists for the current user, and a previous attempt to create one was not authorized.
   */
    SocialStatusCode[SocialStatusCode["NotAuthorizedToCreatePersonalSite"] = 11] = "NotAuthorizedToCreatePersonalSite";
    /**
   * No personal site exists for the current user, and no attempt should be made to create one.
   */
    SocialStatusCode[SocialStatusCode["CannotCreatePersonalSite"] = 12] = "CannotCreatePersonalSite";
    /**
   * The operation was rejected because an internal limit had been reached.
   */
    SocialStatusCode[SocialStatusCode["LimitReached"] = 13] = "LimitReached";
    /**
   * The operation failed because an error occurred during the processing of the specified attachment.
   */
    SocialStatusCode[SocialStatusCode["AttachmentError"] = 14] = "AttachmentError";
    /**
   * The operation succeeded with recoverable errors; the returned data is incomplete.
   */
    SocialStatusCode[SocialStatusCode["PartialData"] = 15] = "PartialData";
    /**
   * A required SharePoint feature is not enabled.
   */
    SocialStatusCode[SocialStatusCode["FeatureDisabled"] = 16] = "FeatureDisabled";
    /**
   * The site's storage quota has been exceeded.
   */
    SocialStatusCode[SocialStatusCode["StorageQuotaExceeded"] = 17] = "StorageQuotaExceeded";
    /**
   * The operation failed because the server could not access the database.
   */
    SocialStatusCode[SocialStatusCode["DatabaseError"] = 18] = "DatabaseError";
})(SocialStatusCode || (SocialStatusCode = {}));


/***/ }),

/***/ 2678:
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/spqueryable.js ***!
  \*********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   SPCollection: () => (/* binding */ SPCollection),
/* harmony export */   SPInstance: () => (/* binding */ SPInstance),
/* harmony export */   SPQueryable: () => (/* binding */ SPQueryable),
/* harmony export */   _SPCollection: () => (/* binding */ _SPCollection),
/* harmony export */   _SPInstance: () => (/* binding */ _SPInstance),
/* harmony export */   _SPQueryable: () => (/* binding */ _SPQueryable),
/* harmony export */   deleteable: () => (/* binding */ deleteable),
/* harmony export */   deleteableWithETag: () => (/* binding */ deleteableWithETag),
/* harmony export */   spDelete: () => (/* binding */ spDelete),
/* harmony export */   spGet: () => (/* binding */ spGet),
/* harmony export */   spInvokableFactory: () => (/* binding */ spInvokableFactory),
/* harmony export */   spPatch: () => (/* binding */ spPatch),
/* harmony export */   spPost: () => (/* binding */ spPost),
/* harmony export */   spPostMerge: () => (/* binding */ spPostMerge)
/* harmony export */ });
/* unused harmony exports spPostDelete, spPostDeleteETag */
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/queryable */ 6844);


const spInvokableFactory = (f) => {
    return (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.queryableFactory)(f);
};
/**
 * SharePointQueryable Base Class
 *
 */
class _SPQueryable extends _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.Queryable {
    /**
     * Creates a new instance of the SharePointQueryable class
     *
     * @constructor
     * @param base A string or SharePointQueryable that should form the base part of the url
     *
     */
    constructor(base, path) {
        if (typeof base === "string") {
            let url = "";
            let parentUrl = "";
            // we need to do some extra parsing to get the parent url correct if we are
            // being created from just a string.
            if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isUrlAbsolute)(base) || base.lastIndexOf("/") < 0) {
                parentUrl = base;
                url = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.combine)(base, path);
            }
            else if (base.lastIndexOf("/") > base.lastIndexOf("(")) {
                // .../items(19)/fields
                const index = base.lastIndexOf("/");
                parentUrl = base.slice(0, index);
                path = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.combine)(base.slice(index), path);
                url = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.combine)(parentUrl, path);
            }
            else {
                // .../items(19)
                const index = base.lastIndexOf("(");
                parentUrl = base.slice(0, index);
                url = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.combine)(base, path);
            }
            // init base with corrected string value
            super(url);
            this.parentUrl = parentUrl;
        }
        else {
            super(base, path);
            const q = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isArray)(base) ? base[0] : base;
            this.parentUrl = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isArray)(base) ? base[1] : q.toUrl();
        }
    }
    /**
     * Gets the full url with query information
     */
    toRequestUrl() {
        const aliasedParams = new URLSearchParams(this.query);
        // this regex is designed to locate aliased parameters within url paths
        let url = this.toUrl().replace(/'!(@.+?)::((?:[^']|'')+)'/ig, (match, labelName, value) => {
            this.log(`Rewriting aliased parameter from match ${match} to label: ${labelName} value: ${value}`, 0);
            aliasedParams.set(labelName, `'${value}'`);
            return labelName;
        });
        const query = aliasedParams.toString();
        if (!(0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.stringIsNullOrEmpty)(query)) {
            url += `${url.indexOf("?") > -1 ? "&" : "?"}${query}`;
        }
        return url;
    }
    /**
     * Choose which fields to return
     *
     * @param selects One or more fields to return
     */
    select(...selects) {
        if (selects.length > 0) {
            this.query.set("$select", selects.join(","));
        }
        return this;
    }
    /**
     * Expands fields such as lookups to get additional data
     *
     * @param expands The Fields for which to expand the values
     */
    expand(...expands) {
        if (expands.length > 0) {
            this.query.set("$expand", expands.join(","));
        }
        return this;
    }
    /**
     * Gets a parent for this instance as specified
     *
     * @param factory The contructor for the class to create
     */
    getParent(factory, path, base = this.parentUrl) {
        return factory([this, base], path);
    }
}
const SPQueryable = spInvokableFactory(_SPQueryable);
/**
 * Represents a REST collection which can be filtered, paged, and selected
 *
 */
class _SPCollection extends _SPQueryable {
    /**
     * Filters the returned collection (https://msdn.microsoft.com/en-us/library/office/fp142385.aspx#bk_supported)
     *
     * @param filter The string representing the filter query
     */
    filter(filter) {
        if (typeof filter === "object") {
            this.query.set("$filter", filter.toString());
            return this;
        }
        if (typeof filter === "function") {
            this.query.set("$filter", filter(SPOData.Where()).toString());
            return this;
        }
        this.query.set("$filter", filter.toString());
        return this;
    }
    /**
     * Orders based on the supplied fields
     *
     * @param orderby The name of the field on which to sort
     * @param ascending If false DESC is appended, otherwise ASC (default)
     */
    orderBy(orderBy, ascending = true) {
        const o = "$orderby";
        const query = this.query.has(o) ? this.query.get(o).split(",") : [];
        query.push(`${orderBy} ${ascending ? "asc" : "desc"}`);
        this.query.set(o, query.join(","));
        return this;
    }
    /**
     * Skips the specified number of items
     *
     * @param skip The number of items to skip
     */
    skip(skip) {
        this.query.set("$skip", skip.toString());
        return this;
    }
    /**
     * Limits the query to only return the specified number of items
     *
     * @param top The query row limit
     */
    top(top) {
        this.query.set("$top", top.toString());
        return this;
    }
}
const SPCollection = spInvokableFactory(_SPCollection);
/**
 * Represents an instance that can be selected
 *
 */
class _SPInstance extends _SPQueryable {
}
const SPInstance = spInvokableFactory(_SPInstance);
/**
 * Adds the a delete method to the tagged class taking no parameters and calling spPostDelete
 */
function deleteable() {
    return function () {
        return spPostDelete(this);
    };
}
function deleteableWithETag() {
    return function (eTag = "*") {
        return spPostDeleteETag(this, {}, eTag);
    };
}
const spGet = (o, init) => {
    return (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.op)(o, _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.get, init);
};
const spPost = (o, init) => (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.op)(o, _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.post, init);
const spPostMerge = (o, init) => {
    init = init || {};
    init.headers = { ...init.headers, "X-HTTP-Method": "MERGE" };
    return spPost(o, init);
};
const spPostDelete = (o, init) => {
    init = init || {};
    init.headers = { ...init.headers || {}, "X-HTTP-Method": "DELETE" };
    return spPost(o, init);
};
const spPostDeleteETag = (o, init, eTag = "*") => {
    init = init || {};
    init.headers = { ...init.headers || {}, "IF-Match": eTag };
    return spPostDelete(o, init);
};
const spDelete = (o, init) => (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.op)(o, _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.del, init);
const spPatch = (o, init) => (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.op)(o, _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.patch, init);
var FilterOperation;
(function (FilterOperation) {
    FilterOperation["Equals"] = "eq";
    FilterOperation["NotEquals"] = "ne";
    FilterOperation["GreaterThan"] = "gt";
    FilterOperation["GreaterThanOrEqualTo"] = "ge";
    FilterOperation["LessThan"] = "lt";
    FilterOperation["LessThanOrEqualTo"] = "le";
    FilterOperation["StartsWith"] = "startswith";
    FilterOperation["SubstringOf"] = "substringof";
})(FilterOperation || (FilterOperation = {}));
var FilterJoinOperator;
(function (FilterJoinOperator) {
    FilterJoinOperator["And"] = "and";
    FilterJoinOperator["AndWithSpace"] = " and ";
    FilterJoinOperator["Or"] = "or";
    FilterJoinOperator["OrWithSpace"] = " or ";
})(FilterJoinOperator || (FilterJoinOperator = {}));
class SPOData {
    static Where() {
        return new InitialFieldQuery([]);
    }
}
// Linting complains that TBaseInterface is unused, but without it all the intellisense is lost since it's carrying it through the chain
class BaseQuery {
    constructor(query) {
        this.query = [];
        this.query = query;
    }
}
class QueryableFields extends BaseQuery {
    constructor(q) {
        super(q);
    }
    text(internalName) {
        return new TextField([...this.query, internalName]);
    }
    choice(internalName) {
        return new TextField([...this.query, internalName]);
    }
    multiChoice(internalName) {
        return new TextField([...this.query, internalName]);
    }
    number(internalName) {
        return new NumberField([...this.query, internalName]);
    }
    date(internalName) {
        return new DateField([...this.query, internalName]);
    }
    boolean(internalName) {
        return new BooleanField([...this.query, internalName]);
    }
    lookup(internalName) {
        return new LookupQueryableFields([...this.query], internalName);
    }
    lookupId(internalName) {
        const col = internalName.endsWith("Id") ? internalName : `${internalName}Id`;
        return new NumberField([...this.query, col]);
    }
}
class QueryableAndResult extends QueryableFields {
    or(...queries) {
        return new ComparisonResult([...this.query, `(${queries.map(x => x.toString()).join(FilterJoinOperator.OrWithSpace)})`]);
    }
}
class QueryableOrResult extends QueryableFields {
    and(...queries) {
        return new ComparisonResult([...this.query, `(${queries.map(x => x.toString()).join(FilterJoinOperator.AndWithSpace)})`]);
    }
}
class InitialFieldQuery extends QueryableFields {
    or(...queries) {
        if (queries == null || queries.length === 0) {
            return new QueryableFields([...this.query, FilterJoinOperator.Or]);
        }
        return new ComparisonResult([...this.query, `(${queries.map(x => x.toString()).join(FilterJoinOperator.OrWithSpace)})`]);
    }
    and(...queries) {
        if (queries == null || queries.length === 0) {
            return new QueryableFields([...this.query, FilterJoinOperator.And]);
        }
        return new ComparisonResult([...this.query, `(${queries.map(x => x.toString()).join(FilterJoinOperator.AndWithSpace)})`]);
    }
}
class LookupQueryableFields extends BaseQuery {
    constructor(q, LookupField) {
        super(q);
        this.LookupField = LookupField;
    }
    Id(id) {
        return new ComparisonResult([...this.query, `${this.LookupField}/Id`, FilterOperation.Equals, id.toString()]);
    }
    text(internalName) {
        return new TextField([...this.query, `${this.LookupField}/${internalName}`]);
    }
    number(internalName) {
        return new NumberField([...this.query, `${this.LookupField}/${internalName}`]);
    }
}
class NullableField extends BaseQuery {
    constructor(q) {
        super(q);
        this.LastIndex = q.length - 1;
        this.InternalName = q[this.LastIndex];
    }
    toODataValue(value) {
        return `'${value}'`;
    }
    isNull() {
        return new ComparisonResult([...this.query, FilterOperation.Equals, "null"]);
    }
    isNotNull() {
        return new ComparisonResult([...this.query, FilterOperation.NotEquals, "null"]);
    }
}
class ComparableField extends NullableField {
    equals(value) {
        return new ComparisonResult([...this.query, FilterOperation.Equals, this.toODataValue(value)]);
    }
    notEquals(value) {
        return new ComparisonResult([...this.query, FilterOperation.NotEquals, this.toODataValue(value)]);
    }
    in(...values) {
        return SPOData.Where().or(...values.map(x => this.equals(x)));
    }
    notIn(...values) {
        return SPOData.Where().and(...values.map(x => this.notEquals(x)));
    }
}
class TextField extends ComparableField {
    startsWith(value) {
        const filter = `${FilterOperation.StartsWith}(${this.InternalName}, ${this.toODataValue(value)})`;
        this.query[this.LastIndex] = filter;
        return new ComparisonResult([...this.query]);
    }
    contains(value) {
        const filter = `${FilterOperation.SubstringOf}(${this.toODataValue(value)}, ${this.InternalName})`;
        this.query[this.LastIndex] = filter;
        return new ComparisonResult([...this.query]);
    }
}
class BooleanField extends NullableField {
    toODataValue(value) {
        return `${value == null ? "null" : value ? 1 : 0}`;
    }
    isTrue() {
        return new ComparisonResult([...this.query, FilterOperation.Equals, this.toODataValue(true)]);
    }
    isFalse() {
        return new ComparisonResult([...this.query, FilterOperation.Equals, this.toODataValue(false)]);
    }
    isFalseOrNull() {
        const filter = `(${[
            this.InternalName,
            FilterOperation.Equals,
            this.toODataValue(null),
            FilterJoinOperator.Or,
            this.InternalName,
            FilterOperation.Equals,
            this.toODataValue(false),
        ].join(" ")})`;
        this.query[this.LastIndex] = filter;
        return new ComparisonResult([...this.query]);
    }
}
class NumericField extends ComparableField {
    greaterThan(value) {
        return new ComparisonResult([...this.query, FilterOperation.GreaterThan, this.toODataValue(value)]);
    }
    greaterThanOrEquals(value) {
        return new ComparisonResult([...this.query, FilterOperation.GreaterThanOrEqualTo, this.toODataValue(value)]);
    }
    lessThan(value) {
        return new ComparisonResult([...this.query, FilterOperation.LessThan, this.toODataValue(value)]);
    }
    lessThanOrEquals(value) {
        return new ComparisonResult([...this.query, FilterOperation.LessThanOrEqualTo, this.toODataValue(value)]);
    }
}
class NumberField extends NumericField {
    toODataValue(value) {
        return `${value}`;
    }
}
class DateField extends NumericField {
    toODataValue(value) {
        return `'${value.toISOString()}'`;
    }
    isBetween(startDate, endDate) {
        const filter = `(${[
            this.InternalName,
            FilterOperation.GreaterThan,
            this.toODataValue(startDate),
            FilterJoinOperator.And,
            this.InternalName,
            FilterOperation.LessThan,
            this.toODataValue(endDate),
        ].join(" ")})`;
        this.query[this.LastIndex] = filter;
        return new ComparisonResult([...this.query]);
    }
    isToday() {
        const StartToday = new Date();
        StartToday.setHours(0, 0, 0, 0);
        const EndToday = new Date();
        EndToday.setHours(23, 59, 59, 999);
        return this.isBetween(StartToday, EndToday);
    }
}
class ComparisonResult extends BaseQuery {
    // eslint-disable-next-line max-len
    and(...queries) {
        if (queries == null || queries.length === 0) {
            return new QueryableAndResult([...this.query, FilterJoinOperator.And]);
        }
        return new ComparisonResult([...this.query, FilterJoinOperator.And, `(${queries.map(x => x.toString()).join(FilterJoinOperator.AndWithSpace)})`]);
    }
    // eslint-disable-next-line max-len
    or(...queries) {
        if (queries == null || queries.length === 0) {
            return new QueryableOrResult([...this.query, FilterJoinOperator.Or]);
        }
        return new ComparisonResult([...this.query, FilterJoinOperator.Or, `(${queries.map(x => x.toString()).join(FilterJoinOperator.OrWithSpace)})`]);
    }
    toString() {
        return this.query.join(" ");
    }
}


/***/ }),

/***/ 4846:
/*!***************************************************!*\
  !*** ./node_modules/@pnp/sp/sputilities/index.js ***!
  \***************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _fi_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../fi.js */ 6553);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 8910);



Reflect.defineProperty(_fi_js__WEBPACK_IMPORTED_MODULE_0__.SPFI.prototype, "utility", {
    configurable: true,
    enumerable: true,
    get: function () {
        return this.create(_types_js__WEBPACK_IMPORTED_MODULE_1__.Utilities, "");
    },
});


/***/ }),

/***/ 8910:
/*!***************************************************!*\
  !*** ./node_modules/@pnp/sp/sputilities/types.js ***!
  \***************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Utilities: () => (/* binding */ Utilities)
/* harmony export */ });
/* unused harmony export _Utilities */
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../utils/extract-web-url.js */ 2997);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/core */ 1971);




class _Utilities extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPQueryable {
    constructor(base, methodName = "") {
        super(base);
        this._url = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_3__.combine)((0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_2__.extractWebUrl)(this._url), `_api/SP.Utilities.Utility.${methodName}`);
    }
    excute(props) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)(props));
    }
    sendEmail(properties) {
        if (properties.AdditionalHeaders) {
            // we have to remap the additional headers into this format #2253
            properties.AdditionalHeaders = Reflect.ownKeys(properties.AdditionalHeaders).map(key => ({
                Key: key,
                Value: Reflect.get(properties.AdditionalHeaders, key),
                ValueType: "Edm.String",
            }));
        }
        return UtilitiesCloneFactory(this, "SendEmail").excute({ properties });
    }
    getCurrentUserEmailAddresses() {
        return UtilitiesCloneFactory(this, "GetCurrentUserEmailAddresses").excute({});
    }
    resolvePrincipal(input, scopes, sources, inputIsEmailOnly, addToUserInfoList, matchUserInfoList = false) {
        const params = {
            addToUserInfoList,
            input,
            inputIsEmailOnly,
            matchUserInfoList,
            scopes,
            sources,
        };
        return UtilitiesCloneFactory(this, "ResolvePrincipalInCurrentContext").excute(params);
    }
    searchPrincipals(input, scopes, sources, groupName, maxCount) {
        const params = {
            groupName: groupName,
            input: input,
            maxCount: maxCount,
            scopes: scopes,
            sources: sources,
        };
        return UtilitiesCloneFactory(this, "SearchPrincipalsUsingContextWeb").excute(params);
    }
    createEmailBodyForInvitation(pageAddress) {
        const params = {
            pageAddress: pageAddress,
        };
        return UtilitiesCloneFactory(this, "CreateEmailBodyForInvitation").excute(params);
    }
    expandGroupsToPrincipals(inputs, maxCount = 30) {
        const params = {
            inputs: inputs,
            maxCount: maxCount,
        };
        const clone = UtilitiesCloneFactory(this, "ExpandGroupsToPrincipals");
        return clone.excute(params);
    }
}
const Utilities = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spInvokableFactory)(_Utilities);
const UtilitiesCloneFactory = (base, path) => Utilities(base, path);


/***/ }),

/***/ 6716:
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/sp/subscriptions/index.js ***!
  \*****************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _list_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./list.js */ 6556);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 4987);




/***/ }),

/***/ 6556:
/*!****************************************************!*\
  !*** ./node_modules/@pnp/sp/subscriptions/list.js ***!
  \****************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _lists_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types.js */ 9706);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 4987);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List, "subscriptions", _types_js__WEBPACK_IMPORTED_MODULE_2__.Subscriptions);


/***/ }),

/***/ 4987:
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/sp/subscriptions/types.js ***!
  \*****************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Subscriptions: () => (/* binding */ Subscriptions)
/* harmony export */ });
/* unused harmony exports _Subscriptions, _Subscription, Subscription */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../decorators.js */ 6540);




let _Subscriptions = class _Subscriptions extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPCollection {
    /**
    * Returns all the webhook subscriptions or the specified webhook subscription
    *
    * @param subscriptionId The id of a specific webhook subscription to retrieve, omit to retrieve all the webhook subscriptions
    */
    getById(subscriptionId) {
        return Subscription(this).concat(`('${subscriptionId}')`);
    }
    /**
     * Creates a new webhook subscription
     *
     * @param notificationUrl The url to receive the notifications
     * @param expirationDate The date and time to expire the subscription in the form YYYY-MM-ddTHH:mm:ss+00:00 (maximum of 6 months)
     * @param clientState A client specific string (optional)
     */
    async add(notificationUrl, expirationDate, clientState) {
        const postBody = {
            "expirationDateTime": expirationDate,
            "notificationUrl": notificationUrl,
            "resource": this.toUrl(),
        };
        if (clientState) {
            postBody.clientState = clientState;
        }
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)(postBody));
    }
};
_Subscriptions = (0,tslib__WEBPACK_IMPORTED_MODULE_2__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_3__.defaultPath)("subscriptions")
], _Subscriptions);

const Subscriptions = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_Subscriptions);
class _Subscription extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPInstance {
    /**
     * Renews this webhook subscription
     *
     * @param expirationDate The date and time to expire the subscription in the form YYYY-MM-ddTHH:mm:ss+00:00 (maximum of 6 months, optional)
     * @param notificationUrl The url to receive the notifications (optional)
     * @param clientState A client specific string (optional)
     */
    async update(expirationDate, notificationUrl, clientState) {
        const postBody = {};
        if (expirationDate) {
            postBody.expirationDateTime = expirationDate;
        }
        if (notificationUrl) {
            postBody.notificationUrl = notificationUrl;
        }
        if (clientState) {
            postBody.clientState = clientState;
        }
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPatch)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)(postBody));
    }
    /**
     * Removes this webhook subscription
     *
     */
    delete() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spDelete)(this);
    }
}
const Subscription = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_Subscription);


/***/ }),

/***/ 8713:
/*!***************************************!*\
  !*** ./node_modules/@pnp/sp/types.js ***!
  \***************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   emptyGuid: () => (/* binding */ emptyGuid)
/* harmony export */ });
/* unused harmony exports PrincipalType, PrincipalSource, PageType */
// reference: https://msdn.microsoft.com/en-us/library/office/dn600183.aspx
const emptyGuid = "00000000-0000-0000-0000-000000000000";
/**
 * Specifies the type of a principal.
 */
var PrincipalType;
(function (PrincipalType) {
    /**
     * Enumeration whose value specifies no principal type.
     */
    PrincipalType[PrincipalType["None"] = 0] = "None";
    /**
     * Enumeration whose value specifies a user as the principal type.
     */
    PrincipalType[PrincipalType["User"] = 1] = "User";
    /**
     * Enumeration whose value specifies a distribution list as the principal type.
     */
    PrincipalType[PrincipalType["DistributionList"] = 2] = "DistributionList";
    /**
     * Enumeration whose value specifies a security group as the principal type.
     */
    PrincipalType[PrincipalType["SecurityGroup"] = 4] = "SecurityGroup";
    /**
     * Enumeration whose value specifies a group as the principal type.
     */
    PrincipalType[PrincipalType["SharePointGroup"] = 8] = "SharePointGroup";
    /**
     * Enumeration whose value specifies all principal types.
     */
    // eslint-disable-next-line no-bitwise
    PrincipalType[PrincipalType["All"] = 15] = "All";
})(PrincipalType || (PrincipalType = {}));
/**
 * Specifies the source of a principal.
 */
var PrincipalSource;
(function (PrincipalSource) {
    /**
     * Enumeration whose value specifies no principal source.
     */
    PrincipalSource[PrincipalSource["None"] = 0] = "None";
    /**
     * Enumeration whose value specifies user information list as the principal source.
     */
    PrincipalSource[PrincipalSource["UserInfoList"] = 1] = "UserInfoList";
    /**
     * Enumeration whose value specifies Active Directory as the principal source.
     */
    PrincipalSource[PrincipalSource["Windows"] = 2] = "Windows";
    /**
     * Enumeration whose value specifies the current membership provider as the principal source.
     */
    PrincipalSource[PrincipalSource["MembershipProvider"] = 4] = "MembershipProvider";
    /**
     * Enumeration whose value specifies the current role provider as the principal source.
     */
    PrincipalSource[PrincipalSource["RoleProvider"] = 8] = "RoleProvider";
    /**
     * Enumeration whose value specifies all principal sources.
     */
    // eslint-disable-next-line no-bitwise
    PrincipalSource[PrincipalSource["All"] = 15] = "All";
})(PrincipalSource || (PrincipalSource = {}));
var PageType;
(function (PageType) {
    PageType[PageType["Invalid"] = -1] = "Invalid";
    PageType[PageType["DefaultView"] = 0] = "DefaultView";
    PageType[PageType["NormalView"] = 1] = "NormalView";
    PageType[PageType["DialogView"] = 2] = "DialogView";
    PageType[PageType["View"] = 3] = "View";
    PageType[PageType["DisplayForm"] = 4] = "DisplayForm";
    PageType[PageType["DisplayFormDialog"] = 5] = "DisplayFormDialog";
    PageType[PageType["EditForm"] = 6] = "EditForm";
    PageType[PageType["EditFormDialog"] = 7] = "EditFormDialog";
    PageType[PageType["NewForm"] = 8] = "NewForm";
    PageType[PageType["NewFormDialog"] = 9] = "NewFormDialog";
    PageType[PageType["SolutionForm"] = 10] = "SolutionForm";
    PageType[PageType["PAGE_MAXITEMS"] = 11] = "PAGE_MAXITEMS";
})(PageType || (PageType = {}));


/***/ }),

/***/ 4655:
/*!***********************************************************!*\
  !*** ./node_modules/@pnp/sp/user-custom-actions/index.js ***!
  \***********************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _list_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./list.js */ 9118);
/* harmony import */ var _web_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./web.js */ 1047);
/* harmony import */ var _site_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./site.js */ 8670);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./types.js */ 6687);






/***/ }),

/***/ 9118:
/*!**********************************************************!*\
  !*** ./node_modules/@pnp/sp/user-custom-actions/list.js ***!
  \**********************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _lists_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types.js */ 9706);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 6687);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List, "userCustomActions", _types_js__WEBPACK_IMPORTED_MODULE_2__.UserCustomActions);


/***/ }),

/***/ 8670:
/*!**********************************************************!*\
  !*** ./node_modules/@pnp/sp/user-custom-actions/site.js ***!
  \**********************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _sites_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sites/types.js */ 1114);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 6687);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_sites_types_js__WEBPACK_IMPORTED_MODULE_1__._Site, "userCustomActions", _types_js__WEBPACK_IMPORTED_MODULE_2__.UserCustomActions);


/***/ }),

/***/ 6687:
/*!***********************************************************!*\
  !*** ./node_modules/@pnp/sp/user-custom-actions/types.js ***!
  \***********************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   UserCustomActions: () => (/* binding */ UserCustomActions),
/* harmony export */   _UserCustomAction: () => (/* binding */ _UserCustomAction)
/* harmony export */ });
/* unused harmony exports _UserCustomActions, UserCustomAction, UserCustomActionRegistrationType, UserCustomActionScope */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../decorators.js */ 6540);




let _UserCustomActions = class _UserCustomActions extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPCollection {
    /**
     * Returns the user custom action with the specified id
     *
     * @param id The GUID id of the user custom action to retrieve
     */
    getById(id) {
        return UserCustomAction(this).concat(`('${id}')`);
    }
    /**
     * Creates a user custom action
     *
     * @param properties The information object of property names and values which define the new user custom action
     */
    async add(properties) {
        const data = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)(properties));
        return {
            action: this.getById(data.Id),
            data,
        };
    }
    /**
     * Deletes all user custom actions in the collection
     */
    clear() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(UserCustomActions(this, "clear"));
    }
};
_UserCustomActions = (0,tslib__WEBPACK_IMPORTED_MODULE_2__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_3__.defaultPath)("usercustomactions")
], _UserCustomActions);

const UserCustomActions = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_UserCustomActions);
class _UserCustomAction extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPInstance {
    constructor() {
        super(...arguments);
        this.delete = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.deleteable)();
    }
    /**
    * Updates this user custom action with the supplied properties
    *
    * @param properties An information object of property names and values to update for this user custom action
    */
    async update(props) {
        const data = await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPostMerge)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)(props));
        return {
            data,
            action: this,
        };
    }
}
const UserCustomAction = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_UserCustomAction);
var UserCustomActionRegistrationType;
(function (UserCustomActionRegistrationType) {
    UserCustomActionRegistrationType[UserCustomActionRegistrationType["None"] = 0] = "None";
    UserCustomActionRegistrationType[UserCustomActionRegistrationType["List"] = 1] = "List";
    UserCustomActionRegistrationType[UserCustomActionRegistrationType["ContentType"] = 2] = "ContentType";
    UserCustomActionRegistrationType[UserCustomActionRegistrationType["ProgId"] = 3] = "ProgId";
    UserCustomActionRegistrationType[UserCustomActionRegistrationType["FileType"] = 4] = "FileType";
})(UserCustomActionRegistrationType || (UserCustomActionRegistrationType = {}));
var UserCustomActionScope;
(function (UserCustomActionScope) {
    UserCustomActionScope[UserCustomActionScope["Unknown"] = 0] = "Unknown";
    UserCustomActionScope[UserCustomActionScope["Site"] = 2] = "Site";
    UserCustomActionScope[UserCustomActionScope["Web"] = 3] = "Web";
    UserCustomActionScope[UserCustomActionScope["List"] = 4] = "List";
})(UserCustomActionScope || (UserCustomActionScope = {}));


/***/ }),

/***/ 1047:
/*!*********************************************************!*\
  !*** ./node_modules/@pnp/sp/user-custom-actions/web.js ***!
  \*********************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _webs_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types.js */ 3169);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 6687);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_webs_types_js__WEBPACK_IMPORTED_MODULE_1__._Web, "userCustomActions", _types_js__WEBPACK_IMPORTED_MODULE_2__.UserCustomActions);


/***/ }),

/***/ 4729:
/*!*******************************************************!*\
  !*** ./node_modules/@pnp/sp/utils/encode-path-str.js ***!
  \*******************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   encodePath: () => (/* binding */ encodePath)
/* harmony export */ });
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);

/**
 * Encodes path portions of SharePoint urls such as decodedUrl=`encodePath(pathStr)`
 *
 * @param value The string path to encode
 * @returns A path encoded for use in SP urls
 */
function encodePath(value) {
    if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.stringIsNullOrEmpty)(value)) {
        return "";
    }
    // replace all instance of ' with ''
    if (/!(@.*?)::(.*?)/ig.test(value)) {
        return value.replace(/!(@.*?)::(.*)$/ig, (match, labelName, v) => {
            // we do not need to encodeURIComponent v as it will be encoded automatically when it is added as a query string param
            // we do need to double any ' chars
            return `!${labelName}::${v.replace(/'/ig, "''")}`;
        });
    }
    else {
        // because this is a literal path value we encodeURIComponent after doubling any ' chars
        return encodeURIComponent(value.replace(/'/ig, "''"));
    }
}


/***/ }),

/***/ 2997:
/*!*******************************************************!*\
  !*** ./node_modules/@pnp/sp/utils/extract-web-url.js ***!
  \*******************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   extractWebUrl: () => (/* binding */ extractWebUrl)
/* harmony export */ });
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);

function extractWebUrl(candidateUrl) {
    if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.stringIsNullOrEmpty)(candidateUrl)) {
        return "";
    }
    let index = candidateUrl.indexOf("_api/");
    if (index < 0) {
        index = candidateUrl.indexOf("_vti_bin/");
    }
    if (index > -1) {
        return candidateUrl.substring(0, index);
    }
    // if all else fails just give them what they gave us back
    return candidateUrl;
}


/***/ }),

/***/ 3999:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/utils/metadata.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   metadata: () => (/* binding */ metadata)
/* harmony export */ });
function metadata(type) {
    return {
        "__metadata": { "type": type },
    };
}


/***/ }),

/***/ 905:
/*!******************************************************!*\
  !*** ./node_modules/@pnp/sp/utils/odata-url-from.js ***!
  \******************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   odataUrlFrom: () => (/* binding */ odataUrlFrom)
/* harmony export */ });
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _extract_web_url_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./extract-web-url.js */ 2997);


function odataUrlFrom(candidate) {
    const parts = [];
    const s = ["odata.type", "odata.editLink", "__metadata", "odata.metadata", "odata.id"];
    if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(candidate, s[0]) && candidate[s[0]] === "SP.Web") {
        // webs return an absolute url in the id
        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(candidate, s[4])) {
            parts.push(candidate[s[4]]);
        }
        else if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(candidate, s[2])) {
            // we are dealing with verbose, which has an absolute uri
            parts.push(candidate.__metadata.uri);
        }
    }
    else {
        if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(candidate, s[3]) && (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(candidate, s[1])) {
            // we are dealign with minimal metadata (default)
            // some entities return an abosolute url in the editlink while for others it is relative
            // without the _api. This code is meant to handle both situations
            const editLink = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.isUrlAbsolute)(candidate[s[1]]) ? candidate[s[1]].split("_api")[1] : candidate[s[1]];
            parts.push((0,_extract_web_url_js__WEBPACK_IMPORTED_MODULE_1__.extractWebUrl)(candidate[s[3]]), "_api", editLink);
        }
        else if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(candidate, s[1])) {
            parts.push("_api", candidate[s[1]]);
        }
        else if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.hOP)(candidate, s[2])) {
            // we are dealing with verbose, which has an absolute uri
            parts.push(candidate.__metadata.uri);
        }
    }
    if (parts.length < 1) {
        return "";
    }
    return (0,_pnp_core__WEBPACK_IMPORTED_MODULE_0__.combine)(...parts);
}


/***/ }),

/***/ 4259:
/*!********************************************************!*\
  !*** ./node_modules/@pnp/sp/utils/to-resource-path.js ***!
  \********************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   toResourcePath: () => (/* binding */ toResourcePath)
/* harmony export */ });
function toResourcePath(url) {
    return {
        DecodedUrl: url,
    };
}


/***/ }),

/***/ 459:
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/views/index.js ***!
  \*********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _list_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./list.js */ 7309);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 5696);




/***/ }),

/***/ 7309:
/*!********************************************!*\
  !*** ./node_modules/@pnp/sp/views/list.js ***!
  \********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _lists_types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types.js */ 9706);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types.js */ 5696);



(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List, "views", _types_js__WEBPACK_IMPORTED_MODULE_2__.Views);
(0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.addProp)(_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List, "defaultView", _types_js__WEBPACK_IMPORTED_MODULE_2__.View);
_lists_types_js__WEBPACK_IMPORTED_MODULE_1__._List.prototype.getView = function (viewId) {
    return (0,_types_js__WEBPACK_IMPORTED_MODULE_2__.View)(this, `getView('${viewId}')`);
};


/***/ }),

/***/ 5696:
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/views/types.js ***!
  \*********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   View: () => (/* binding */ View),
/* harmony export */   Views: () => (/* binding */ Views)
/* harmony export */ });
/* unused harmony exports _Views, _View, _ViewFields, ViewFields, ViewScope */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../decorators.js */ 6540);
/* harmony import */ var _utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../utils/encode-path-str.js */ 4729);





let _Views = class _Views extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPCollection {
    /**
     * Adds a new view to the collection
     *
     * @param title The new views's title
     * @param personalView True if this is a personal view, otherwise false, default = false
     * @param additionalSettings Will be passed as part of the view creation body
     */
    async add(Title, PersonalView = false, additionalSettings = {}) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({
            PersonalView,
            Title,
            ...additionalSettings,
        }));
    }
    /**
     * Gets a view by guid id
     *
     * @param id The GUID id of the view
     */
    getById(id) {
        return View(this).concat(`('${id}')`);
    }
    /**
     * Gets a view by title (case-sensitive)
     *
     * @param title The case-sensitive title of the view
     */
    getByTitle(title) {
        return View(this, `getByTitle('${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_2__.encodePath)(title)}')`);
    }
};
_Views = (0,tslib__WEBPACK_IMPORTED_MODULE_3__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_4__.defaultPath)("views")
], _Views);

const Views = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spInvokableFactory)(_Views);
class _View extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPInstance {
    constructor() {
        super(...arguments);
        this.delete = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.deleteable)();
    }
    get fields() {
        return ViewFields(this);
    }
    /**
     * Updates this view intance with the supplied properties
     *
     * @param properties A plain object hash of values to update for the view
     */
    async update(props) {
        return await (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPostMerge)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)(props));
    }
    // : any = this._update<IViewUpdateResult, ITypedHash<any>>("SP.View", data => ({ data, view: <any>this }));
    /**
     * Returns the list view as HTML.
     *
     */
    renderAsHtml() {
        return View(this, "renderashtml")();
    }
    /**
     * Sets the view schema
     *
     * @param viewXml The view XML to set
     */
    setViewXml(viewXml) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(View(this, "SetViewXml"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({ viewXml }));
    }
}
const View = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spInvokableFactory)(_View);
let _ViewFields = class _ViewFields extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPCollection {
    /**
     * Gets a value that specifies the XML schema that represents the collection.
     */
    getSchemaXml() {
        return ViewFields(this, "schemaxml")();
    }
    /**
     * Adds the field with the specified field internal name or display name to the collection.
     *
     * @param fieldTitleOrInternalName The case-sensitive internal name or display name of the field to add.
     */
    add(fieldTitleOrInternalName) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(ViewFields(this, `addviewfield('${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_2__.encodePath)(fieldTitleOrInternalName)}')`));
    }
    /**
     * Moves the field with the specified field internal name to the specified position in the collection.
     *
     * @param field The case-sensitive internal name of the field to move.
     * @param index The zero-based index of the new position for the field.
     */
    move(field, index) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(ViewFields(this, "moveviewfieldto"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({ field, index }));
    }
    /**
     * Removes all the fields from the collection.
     */
    removeAll() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(ViewFields(this, "removeallviewfields"));
    }
    /**
     * Removes the field with the specified field internal name from the collection.
     *
     * @param fieldInternalName The case-sensitive internal name of the field to remove from the view.
     */
    remove(fieldInternalName) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(ViewFields(this, `removeviewfield('${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_2__.encodePath)(fieldInternalName)}')`));
    }
};
_ViewFields = (0,tslib__WEBPACK_IMPORTED_MODULE_3__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_4__.defaultPath)("viewfields")
], _ViewFields);

const ViewFields = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spInvokableFactory)(_ViewFields);
var ViewScope;
(function (ViewScope) {
    ViewScope[ViewScope["DefaultValue"] = 0] = "DefaultValue";
    ViewScope[ViewScope["Recursive"] = 1] = "Recursive";
    ViewScope[ViewScope["RecursiveAll"] = 2] = "RecursiveAll";
    ViewScope[ViewScope["FilesOnly"] = 3] = "FilesOnly";
})(ViewScope || (ViewScope = {}));


/***/ }),

/***/ 959:
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/webparts/file.js ***!
  \***********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _files_types_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../files/types.js */ 242);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 4294);


_files_types_js__WEBPACK_IMPORTED_MODULE_0__._File.prototype.getLimitedWebPartManager = function (scope = _types_js__WEBPACK_IMPORTED_MODULE_1__.WebPartsPersonalizationScope.Shared) {
    return (0,_types_js__WEBPACK_IMPORTED_MODULE_1__.LimitedWebPartManager)(this, `getLimitedWebPartManager(scope=${scope})`);
};


/***/ }),

/***/ 7765:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/webparts/index.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _file_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./file.js */ 959);
/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types.js */ 4294);




/***/ }),

/***/ 4294:
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/webparts/types.js ***!
  \************************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   LimitedWebPartManager: () => (/* binding */ LimitedWebPartManager),
/* harmony export */   WebPartsPersonalizationScope: () => (/* binding */ WebPartsPersonalizationScope)
/* harmony export */ });
/* unused harmony exports _LimitedWebPartManager, _WebPartDefinitions, WebPartDefinitions, _WebPartDefinition, WebPartDefinition */
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/queryable */ 6844);


class _LimitedWebPartManager extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPQueryable {
    get scope() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.SPQueryable)(this, "Scope");
    }
    get webparts() {
        return WebPartDefinitions(this, "webparts");
    }
    export(id) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(LimitedWebPartManagerCloneFactory(this, "ExportWebPart"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)({ webPartId: id }));
    }
    import(xml) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(LimitedWebPartManagerCloneFactory(this, "ImportWebPart"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_1__.body)({ webPartXml: xml }));
    }
}
const LimitedWebPartManager = (baseUrl, path) => new _LimitedWebPartManager(baseUrl, path);
const LimitedWebPartManagerCloneFactory = (baseUrl, path) => LimitedWebPartManager(baseUrl, path);
class _WebPartDefinitions extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPCollection {
    /**
     * Gets a web part definition from the collection by id
     *
     * @param id The storage ID of the SPWebPartDefinition to retrieve
     */
    getById(id) {
        return WebPartDefinition(this, `getbyid('${id}')`);
    }
    /**
     * Gets a web part definition from the collection by storage id
     *
     * @param id The WebPart.ID of the SPWebPartDefinition to retrieve
     */
    getByControlId(id) {
        return WebPartDefinition(this, `getByControlId('${id}')`);
    }
}
const WebPartDefinitions = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_WebPartDefinitions);
class _WebPartDefinition extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_0__._SPInstance {
    /**
    * Gets the webpart information associated with this definition
    */
    get webpart() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.SPInstance)(this, "webpart");
    }
    /**
     * Saves changes to the Web Part made using other properties and methods on the SPWebPartDefinition object
     */
    saveChanges() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(WebPartDefinition(this, "SaveWebPartChanges"));
    }
    /**
     * Moves the Web Part to a different location on a Web Part Page
     *
     * @param zoneId The ID of the Web Part Zone to which to move the Web Part
     * @param zoneIndex A Web Part zone index that specifies the position at which the Web Part is to be moved within the destination Web Part zone
     */
    moveTo(zoneId, zoneIndex) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(WebPartDefinition(this, `MoveWebPartTo(zoneID='${zoneId}', zoneIndex=${zoneIndex})`));
    }
    /**
     * Closes the Web Part. If the Web Part is already closed, this method does nothing
     */
    close() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(WebPartDefinition(this, "CloseWebPart"));
    }
    /**
     * Opens the Web Part. If the Web Part is already closed, this method does nothing
     */
    open() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(WebPartDefinition(this, "OpenWebPart"));
    }
    /**
     * Removes a webpart from a page, all settings will be lost
     */
    delete() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spPost)(WebPartDefinition(this, "DeleteWebPart"));
    }
}
const WebPartDefinition = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_0__.spInvokableFactory)(_WebPartDefinition);
var WebPartsPersonalizationScope;
(function (WebPartsPersonalizationScope) {
    WebPartsPersonalizationScope[WebPartsPersonalizationScope["User"] = 0] = "User";
    WebPartsPersonalizationScope[WebPartsPersonalizationScope["Shared"] = 1] = "Shared";
})(WebPartsPersonalizationScope || (WebPartsPersonalizationScope = {}));


/***/ }),

/***/ 3867:
/*!********************************************!*\
  !*** ./node_modules/@pnp/sp/webs/index.js ***!
  \********************************************/
/***/ ((__unused_webpack___webpack_module__, __unused_webpack___webpack_exports__, __webpack_require__) => {

/* harmony import */ var _types_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./types.js */ 3169);
/* harmony import */ var _fi_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../fi.js */ 6553);



Reflect.defineProperty(_fi_js__WEBPACK_IMPORTED_MODULE_1__.SPFI.prototype, "web", {
    configurable: true,
    enumerable: true,
    get: function () {
        return this.create(_types_js__WEBPACK_IMPORTED_MODULE_0__.Web);
    },
});


/***/ }),

/***/ 3169:
/*!********************************************!*\
  !*** ./node_modules/@pnp/sp/webs/types.js ***!
  \********************************************/
/***/ ((__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Web: () => (/* binding */ Web),
/* harmony export */   _Web: () => (/* binding */ _Web)
/* harmony export */ });
/* unused harmony exports _Webs, Webs */
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! tslib */ 4346);
/* harmony import */ var _pnp_queryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/queryable */ 6844);
/* harmony import */ var _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../spqueryable.js */ 2678);
/* harmony import */ var _decorators_js__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../decorators.js */ 6540);
/* harmony import */ var _utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../utils/extract-web-url.js */ 2997);
/* harmony import */ var _pnp_core__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/core */ 1971);
/* harmony import */ var _utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../utils/encode-path-str.js */ 4729);







let _Webs = class _Webs extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPCollection {
    /**
     * Adds a new web to the collection
     *
     * @param title The new web's title
     * @param url The new web's relative url
     * @param description The new web's description
     * @param template The new web's template internal name (default = STS)
     * @param language The locale id that specifies the new web's language (default = 1033 [English, US])
     * @param inheritPermissions When true, permissions will be inherited from the new web's parent (default = true)
     */
    async add(Title, Url, Description = "", WebTemplate = "STS", Language = 1033, UseSamePermissionsAsParentSite = true) {
        const postBody = (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({
            "parameters": {
                Description,
                Language,
                Title,
                Url,
                UseSamePermissionsAsParentSite,
                WebTemplate,
            },
        });
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(Webs(this, "add"), postBody);
    }
};
_Webs = (0,tslib__WEBPACK_IMPORTED_MODULE_5__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_6__.defaultPath)("webs")
], _Webs);

const Webs = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spInvokableFactory)(_Webs);
/**
 * Ensures the url passed to the constructor is correctly rebased to a web url
 *
 * @param candidate The candidate web url
 * @param path The caller supplied path, which may contain _api, meaning we don't append _api/web
 */
function rebaseWebUrl(candidate, path) {
    let replace = "_api/web";
    // this allows us to both:
    // - test if `candidate` already has an api path
    // - ensure that we append the correct one as sometimes a web is not defined
    //   by _api/web, in the case of _api/site/rootweb for example
    const matches = /(_api[/|\\](site\/rootweb|site|web))/i.exec(candidate);
    if ((matches === null || matches === void 0 ? void 0 : matches.length) > 0) {
        // we want just the base url part (before the _api)
        candidate = (0,_utils_extract_web_url_js__WEBPACK_IMPORTED_MODULE_2__.extractWebUrl)(candidate);
        // we want to ensure we put back the correct string
        replace = matches[1];
    }
    // we only need to append the _api part IF `path` doesn't already include it.
    if ((path === null || path === void 0 ? void 0 : path.indexOf("_api")) < 0) {
        candidate = (0,_pnp_core__WEBPACK_IMPORTED_MODULE_3__.combine)(candidate, replace);
    }
    return candidate;
}
/**
 * Describes a web
 *
 */
let _Web = class _Web extends _spqueryable_js__WEBPACK_IMPORTED_MODULE_1__._SPInstance {
    constructor(base, path) {
        if (typeof base === "string") {
            base = rebaseWebUrl(base, path);
        }
        else if ((0,_pnp_core__WEBPACK_IMPORTED_MODULE_3__.isArray)(base)) {
            base = [base[0], rebaseWebUrl(base[1], path)];
        }
        else {
            base = [base, rebaseWebUrl(base.toUrl(), path)];
        }
        super(base, path);
        this.delete = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.deleteable)();
    }
    /**
     * Gets this web's subwebs
     *
     */
    get webs() {
        return Webs(this);
    }
    /**
     * Allows access to the web's all properties collection
     */
    get allProperties() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.SPInstance)(this, "allproperties");
    }
    /**
     * Gets a collection of WebInfos for this web's subwebs
     *
     */
    get webinfos() {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.SPCollection)(this, "webinfos");
    }
    /**
     * Gets this web's parent web and data
     *
     */
    async getParentWeb() {
        const { Url, ParentWeb } = await this.select("Url", "ParentWeb/ServerRelativeUrl").expand("ParentWeb")();
        if (ParentWeb === null || ParentWeb === void 0 ? void 0 : ParentWeb.ServerRelativeUrl) {
            return Web([this, (0,_pnp_core__WEBPACK_IMPORTED_MODULE_3__.combine)((new URL(Url)).origin, ParentWeb.ServerRelativeUrl)]);
        }
        return null;
    }
    /**
     * Updates this web instance with the supplied properties
     *
     * @param properties A plain object hash of values to update for the web
     */
    async update(properties) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPostMerge)(this, (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)(properties));
    }
    /**
     * Applies the theme specified by the contents of each of the files specified in the arguments to the site
     *
     * @param colorPaletteUrl The server-relative URL of the color palette file
     * @param fontSchemeUrl The server-relative URL of the font scheme
     * @param backgroundImageUrl The server-relative URL of the background image
     * @param shareGenerated When true, the generated theme files are stored in the root site. When false, they are stored in this web
     */
    applyTheme(colorPaletteUrl, fontSchemeUrl, backgroundImageUrl, shareGenerated) {
        const postBody = (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({
            backgroundImageUrl,
            colorPaletteUrl,
            fontSchemeUrl,
            shareGenerated,
        });
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(Web(this, "applytheme"), postBody);
    }
    /**
     * Applies the specified site definition or site template to the Web site that has no template applied to it
     *
     * @param template Name of the site definition or the name of the site template
     */
    applyWebTemplate(template) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(Web(this, `applywebtemplate(webTemplate='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_4__.encodePath)(template)}')`));
    }
    /**
     * Returns the collection of changes from the change log that have occurred within the list, based on the specified query
     *
     * @param query The change query
     */
    getChanges(query) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(Web(this, "getchanges"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({ query }));
    }
    /**
     * Returns the name of the image file for the icon that is used to represent the specified file
     *
     * @param filename The file name. If this parameter is empty, the server returns an empty string
     * @param size The size of the icon: 16x16 pixels = 0, 32x32 pixels = 1 (default = 0)
     * @param progId The ProgID of the application that was used to create the file, in the form OLEServerName.ObjectName
     */
    mapToIcon(filename, size = 0, progId = "") {
        return Web(this, `maptoicon(filename='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_4__.encodePath)(filename)}',progid='${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_4__.encodePath)(progId)}',size=${size})`)();
    }
    /**
     * Returns the tenant property corresponding to the specified key in the app catalog site
     *
     * @param key Id of storage entity to be set
     */
    getStorageEntity(key) {
        return Web(this, `getStorageEntity('${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_4__.encodePath)(key)}')`)();
    }
    /**
     * This will set the storage entity identified by the given key (MUST be called in the context of the app catalog)
     *
     * @param key Id of storage entity to be set
     * @param value Value of storage entity to be set
     * @param description Description of storage entity to be set
     * @param comments Comments of storage entity to be set
     */
    setStorageEntity(key, value, description = "", comments = "") {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(Web(this, "setStorageEntity"), (0,_pnp_queryable__WEBPACK_IMPORTED_MODULE_0__.body)({
            comments,
            description,
            key,
            value,
        }));
    }
    /**
     * This will remove the storage entity identified by the given key
     *
     * @param key Id of storage entity to be removed
     */
    removeStorageEntity(key) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spPost)(Web(this, `removeStorageEntity('${(0,_utils_encode_path_str_js__WEBPACK_IMPORTED_MODULE_4__.encodePath)(key)}')`));
    }
    /**
    * Returns a collection of objects that contain metadata about subsites of the current site in which the current user is a member.
    *
    * @param nWebTemplateFilter Specifies the site definition (default = -1)
    * @param nConfigurationFilter A 16-bit integer that specifies the identifier of a configuration (default = -1)
    */
    getSubwebsFilteredForCurrentUser(nWebTemplateFilter = -1, nConfigurationFilter = -1) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.SPCollection)(this, `getSubwebsFilteredForCurrentUser(nWebTemplateFilter=${nWebTemplateFilter},nConfigurationFilter=${nConfigurationFilter})`);
    }
    /**
     * Returns a collection of site templates available for the site
     *
     * @param language The locale id of the site templates to retrieve (default = 1033 [English, US])
     * @param includeCrossLanguage When true, includes language-neutral site templates; otherwise false (default = true)
     */
    availableWebTemplates(language = 1033, includeCrossLanugage = true) {
        return (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.SPCollection)(this, `getavailablewebtemplates(lcid=${language},doincludecrosslanguage=${includeCrossLanugage})`);
    }
};
_Web = (0,tslib__WEBPACK_IMPORTED_MODULE_5__.__decorate)([
    (0,_decorators_js__WEBPACK_IMPORTED_MODULE_6__.defaultPath)("_api/web")
], _Web);

const Web = (0,_spqueryable_js__WEBPACK_IMPORTED_MODULE_1__.spInvokableFactory)(_Web);


/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/compat get default export */
/******/ 	(() => {
/******/ 		// getDefaultExport function for compatibility with non-harmony modules
/******/ 		__webpack_require__.n = (module) => {
/******/ 			var getter = module && module.__esModule ?
/******/ 				() => (module['default']) :
/******/ 				() => (module);
/******/ 			__webpack_require__.d(getter, { a: getter });
/******/ 			return getter;
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/define property getters */
/******/ 	(() => {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = (exports, definition) => {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	(() => {
/******/ 		__webpack_require__.o = (obj, prop) => (Object.prototype.hasOwnProperty.call(obj, prop))
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	(() => {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = (exports) => {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	})();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry need to be wrapped in an IIFE because it need to be isolated against other modules in the chunk.
(() => {
/*!****************************************************!*\
  !*** ./lib/webparts/holaMundo/HolaMundoWebPart.js ***!
  \****************************************************/
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var _microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @microsoft/sp-webpart-base */ 3134);
/* harmony import */ var _microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _pnp_sp_presets_all__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/sp/presets/all */ 6935);
var __extends = (undefined && undefined.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (undefined && undefined.__generator) || function (thisArg, body) {
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


var HolaMundoWebPart = /** @class */ (function (_super) {
    __extends(HolaMundoWebPart, _super);
    function HolaMundoWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    HolaMundoWebPart.prototype.render = function () {
        return __awaiter(this, void 0, void 0, function () {
            var lists, documentLibraries, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        // Inicializar SPFI con el contexto de SPFx
                        this.sp = (0,_pnp_sp_presets_all__WEBPACK_IMPORTED_MODULE_1__.spfi)().using((0,_pnp_sp_presets_all__WEBPACK_IMPORTED_MODULE_1__.SPFx)(this.context)); // Usa this.context de BaseClientSideWebPart
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        // Obtener todas las listas del sitio actual
                        console.log("Obteniendo listas desde SharePoint...");
                        return [4 /*yield*/, this.sp.web.lists()];
                    case 2:
                        lists = _a.sent();
                        console.log("Listas obtenidas correctamente:", lists);
                        if (!lists) {
                            console.error("No se pudieron obtener las listas.");
                            return [2 /*return*/];
                        }
                        // Contar todas las listas
                        console.log("Número de listas y bibliotecas: ", lists.length);
                        documentLibraries = lists.filter(function (list) { return list.BaseTemplate === 101; });
                        console.log("Número de bibliotecas de documentos: ", documentLibraries.length);
                        return [3 /*break*/, 4];
                    case 3:
                        error_1 = _a.sent();
                        console.error("Error al obtener las listas: ", error_1);
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    return HolaMundoWebPart;
}(_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_0__.BaseClientSideWebPart));
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (HolaMundoWebPart);

})();

/******/ 	return __webpack_exports__;
/******/ })()
;
});;
//# sourceMappingURL=hola-mundo-web-part.js.map