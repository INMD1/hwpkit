var yi = Object.defineProperty;
var Ai = (t, r, e) => r in t ? yi(t, r, { enumerable: !0, configurable: !0, writable: !0, value: e }) : t[r] = e;
var ce = (t, r, e) => Ai(t, typeof r != "symbol" ? r + "" : r, e);
import { useState as Oe, useCallback as Si } from "react";
var mr = typeof globalThis < "u" ? globalThis : typeof window < "u" ? window : typeof global < "u" ? global : typeof self < "u" ? self : {};
function bi(t) {
  return t && t.__esModule && Object.prototype.hasOwnProperty.call(t, "default") ? t.default : t;
}
function xi(t) {
  if (t.__esModule) return t;
  var r = t.default;
  if (typeof r == "function") {
    var e = function n() {
      return this instanceof n ? Reflect.construct(r, arguments, this.constructor) : r.apply(this, arguments);
    };
    e.prototype = r.prototype;
  } else e = {};
  return Object.defineProperty(e, "__esModule", { value: !0 }), Object.keys(t).forEach(function(n) {
    var i = Object.getOwnPropertyDescriptor(t, n);
    Object.defineProperty(e, n, i.get ? i : {
      enumerable: !0,
      get: function() {
        return t[n];
      }
    });
  }), e;
}
function jt(t) {
  throw new Error('Could not dynamically require "' + t + '". Please configure the dynamicRequireTargets or/and ignoreDynamicRequires option of @rollup/plugin-commonjs appropriately for this require call to work.');
}
var ma = { exports: {} };
/*!

JSZip v3.10.1 - A JavaScript class for generating and reading zip files
<http://stuartk.com/jszip>

(c) 2009-2016 Stuart Knightley <stuart [at] stuartk.com>
Dual licenced under the MIT license or GPLv3. See https://raw.github.com/Stuk/jszip/main/LICENSE.markdown.

JSZip uses the library pako released under the MIT license :
https://github.com/nodeca/pako/blob/main/LICENSE
*/
(function(t, r) {
  (function(e) {
    t.exports = e();
  })(function() {
    return function e(n, i, a) {
      function s(l, f) {
        if (!i[l]) {
          if (!n[l]) {
            var d = typeof jt == "function" && jt;
            if (!f && d) return d(l, !0);
            if (o) return o(l, !0);
            var u = new Error("Cannot find module '" + l + "'");
            throw u.code = "MODULE_NOT_FOUND", u;
          }
          var c = i[l] = { exports: {} };
          n[l][0].call(c.exports, function(g) {
            var p = n[l][1][g];
            return s(p || g);
          }, c, c.exports, e, n, i, a);
        }
        return i[l].exports;
      }
      for (var o = typeof jt == "function" && jt, h = 0; h < a.length; h++) s(a[h]);
      return s;
    }({ 1: [function(e, n, i) {
      var a = e("./utils"), s = e("./support"), o = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
      i.encode = function(h) {
        for (var l, f, d, u, c, g, p, y = [], m = 0, A = h.length, E = A, x = a.getTypeOf(h) !== "string"; m < h.length; ) E = A - m, d = x ? (l = h[m++], f = m < A ? h[m++] : 0, m < A ? h[m++] : 0) : (l = h.charCodeAt(m++), f = m < A ? h.charCodeAt(m++) : 0, m < A ? h.charCodeAt(m++) : 0), u = l >> 2, c = (3 & l) << 4 | f >> 4, g = 1 < E ? (15 & f) << 2 | d >> 6 : 64, p = 2 < E ? 63 & d : 64, y.push(o.charAt(u) + o.charAt(c) + o.charAt(g) + o.charAt(p));
        return y.join("");
      }, i.decode = function(h) {
        var l, f, d, u, c, g, p = 0, y = 0, m = "data:";
        if (h.substr(0, m.length) === m) throw new Error("Invalid base64 input, it looks like a data url.");
        var A, E = 3 * (h = h.replace(/[^A-Za-z0-9+/=]/g, "")).length / 4;
        if (h.charAt(h.length - 1) === o.charAt(64) && E--, h.charAt(h.length - 2) === o.charAt(64) && E--, E % 1 != 0) throw new Error("Invalid base64 input, bad content length.");
        for (A = s.uint8array ? new Uint8Array(0 | E) : new Array(0 | E); p < h.length; ) l = o.indexOf(h.charAt(p++)) << 2 | (u = o.indexOf(h.charAt(p++))) >> 4, f = (15 & u) << 4 | (c = o.indexOf(h.charAt(p++))) >> 2, d = (3 & c) << 6 | (g = o.indexOf(h.charAt(p++))), A[y++] = l, c !== 64 && (A[y++] = f), g !== 64 && (A[y++] = d);
        return A;
      };
    }, { "./support": 30, "./utils": 32 }], 2: [function(e, n, i) {
      var a = e("./external"), s = e("./stream/DataWorker"), o = e("./stream/Crc32Probe"), h = e("./stream/DataLengthProbe");
      function l(f, d, u, c, g) {
        this.compressedSize = f, this.uncompressedSize = d, this.crc32 = u, this.compression = c, this.compressedContent = g;
      }
      l.prototype = { getContentWorker: function() {
        var f = new s(a.Promise.resolve(this.compressedContent)).pipe(this.compression.uncompressWorker()).pipe(new h("data_length")), d = this;
        return f.on("end", function() {
          if (this.streamInfo.data_length !== d.uncompressedSize) throw new Error("Bug : uncompressed data size mismatch");
        }), f;
      }, getCompressedWorker: function() {
        return new s(a.Promise.resolve(this.compressedContent)).withStreamInfo("compressedSize", this.compressedSize).withStreamInfo("uncompressedSize", this.uncompressedSize).withStreamInfo("crc32", this.crc32).withStreamInfo("compression", this.compression);
      } }, l.createWorkerFrom = function(f, d, u) {
        return f.pipe(new o()).pipe(new h("uncompressedSize")).pipe(d.compressWorker(u)).pipe(new h("compressedSize")).withStreamInfo("compression", d);
      }, n.exports = l;
    }, { "./external": 6, "./stream/Crc32Probe": 25, "./stream/DataLengthProbe": 26, "./stream/DataWorker": 27 }], 3: [function(e, n, i) {
      var a = e("./stream/GenericWorker");
      i.STORE = { magic: "\0\0", compressWorker: function() {
        return new a("STORE compression");
      }, uncompressWorker: function() {
        return new a("STORE decompression");
      } }, i.DEFLATE = e("./flate");
    }, { "./flate": 7, "./stream/GenericWorker": 28 }], 4: [function(e, n, i) {
      var a = e("./utils"), s = function() {
        for (var o, h = [], l = 0; l < 256; l++) {
          o = l;
          for (var f = 0; f < 8; f++) o = 1 & o ? 3988292384 ^ o >>> 1 : o >>> 1;
          h[l] = o;
        }
        return h;
      }();
      n.exports = function(o, h) {
        return o !== void 0 && o.length ? a.getTypeOf(o) !== "string" ? function(l, f, d, u) {
          var c = s, g = u + d;
          l ^= -1;
          for (var p = u; p < g; p++) l = l >>> 8 ^ c[255 & (l ^ f[p])];
          return -1 ^ l;
        }(0 | h, o, o.length, 0) : function(l, f, d, u) {
          var c = s, g = u + d;
          l ^= -1;
          for (var p = u; p < g; p++) l = l >>> 8 ^ c[255 & (l ^ f.charCodeAt(p))];
          return -1 ^ l;
        }(0 | h, o, o.length, 0) : 0;
      };
    }, { "./utils": 32 }], 5: [function(e, n, i) {
      i.base64 = !1, i.binary = !1, i.dir = !1, i.createFolders = !0, i.date = null, i.compression = null, i.compressionOptions = null, i.comment = null, i.unixPermissions = null, i.dosPermissions = null;
    }, {}], 6: [function(e, n, i) {
      var a = null;
      a = typeof Promise < "u" ? Promise : e("lie"), n.exports = { Promise: a };
    }, { lie: 37 }], 7: [function(e, n, i) {
      var a = typeof Uint8Array < "u" && typeof Uint16Array < "u" && typeof Uint32Array < "u", s = e("pako"), o = e("./utils"), h = e("./stream/GenericWorker"), l = a ? "uint8array" : "array";
      function f(d, u) {
        h.call(this, "FlateWorker/" + d), this._pako = null, this._pakoAction = d, this._pakoOptions = u, this.meta = {};
      }
      i.magic = "\b\0", o.inherits(f, h), f.prototype.processChunk = function(d) {
        this.meta = d.meta, this._pako === null && this._createPako(), this._pako.push(o.transformTo(l, d.data), !1);
      }, f.prototype.flush = function() {
        h.prototype.flush.call(this), this._pako === null && this._createPako(), this._pako.push([], !0);
      }, f.prototype.cleanUp = function() {
        h.prototype.cleanUp.call(this), this._pako = null;
      }, f.prototype._createPako = function() {
        this._pako = new s[this._pakoAction]({ raw: !0, level: this._pakoOptions.level || -1 });
        var d = this;
        this._pako.onData = function(u) {
          d.push({ data: u, meta: d.meta });
        };
      }, i.compressWorker = function(d) {
        return new f("Deflate", d);
      }, i.uncompressWorker = function() {
        return new f("Inflate", {});
      };
    }, { "./stream/GenericWorker": 28, "./utils": 32, pako: 38 }], 8: [function(e, n, i) {
      function a(c, g) {
        var p, y = "";
        for (p = 0; p < g; p++) y += String.fromCharCode(255 & c), c >>>= 8;
        return y;
      }
      function s(c, g, p, y, m, A) {
        var E, x, S = c.file, R = c.compression, v = A !== l.utf8encode, O = o.transformTo("string", A(S.name)), C = o.transformTo("string", l.utf8encode(S.name)), B = S.comment, Z = o.transformTo("string", A(B)), T = o.transformTo("string", l.utf8encode(B)), j = C.length !== S.name.length, _ = T.length !== B.length, K = "", se = "", P = "", G = S.dir, U = S.date, Q = { crc32: 0, compressedSize: 0, uncompressedSize: 0 };
        g && !p || (Q.crc32 = c.crc32, Q.compressedSize = c.compressedSize, Q.uncompressedSize = c.uncompressedSize);
        var M = 0;
        g && (M |= 8), v || !j && !_ || (M |= 2048);
        var X = 0, de = 0;
        G && (X |= 16), m === "UNIX" ? (de = 798, X |= function(ne, ye) {
          var Te = ne;
          return ne || (Te = ye ? 16893 : 33204), (65535 & Te) << 16;
        }(S.unixPermissions, G)) : (de = 20, X |= function(ne) {
          return 63 & (ne || 0);
        }(S.dosPermissions)), E = U.getUTCHours(), E <<= 6, E |= U.getUTCMinutes(), E <<= 5, E |= U.getUTCSeconds() / 2, x = U.getUTCFullYear() - 1980, x <<= 4, x |= U.getUTCMonth() + 1, x <<= 5, x |= U.getUTCDate(), j && (se = a(1, 1) + a(f(O), 4) + C, K += "up" + a(se.length, 2) + se), _ && (P = a(1, 1) + a(f(Z), 4) + T, K += "uc" + a(P.length, 2) + P);
        var le = "";
        return le += `
\0`, le += a(M, 2), le += R.magic, le += a(E, 2), le += a(x, 2), le += a(Q.crc32, 4), le += a(Q.compressedSize, 4), le += a(Q.uncompressedSize, 4), le += a(O.length, 2), le += a(K.length, 2), { fileRecord: d.LOCAL_FILE_HEADER + le + O + K, dirRecord: d.CENTRAL_FILE_HEADER + a(de, 2) + le + a(Z.length, 2) + "\0\0\0\0" + a(X, 4) + a(y, 4) + O + K + Z };
      }
      var o = e("../utils"), h = e("../stream/GenericWorker"), l = e("../utf8"), f = e("../crc32"), d = e("../signature");
      function u(c, g, p, y) {
        h.call(this, "ZipFileWorker"), this.bytesWritten = 0, this.zipComment = g, this.zipPlatform = p, this.encodeFileName = y, this.streamFiles = c, this.accumulate = !1, this.contentBuffer = [], this.dirRecords = [], this.currentSourceOffset = 0, this.entriesCount = 0, this.currentFile = null, this._sources = [];
      }
      o.inherits(u, h), u.prototype.push = function(c) {
        var g = c.meta.percent || 0, p = this.entriesCount, y = this._sources.length;
        this.accumulate ? this.contentBuffer.push(c) : (this.bytesWritten += c.data.length, h.prototype.push.call(this, { data: c.data, meta: { currentFile: this.currentFile, percent: p ? (g + 100 * (p - y - 1)) / p : 100 } }));
      }, u.prototype.openedSource = function(c) {
        this.currentSourceOffset = this.bytesWritten, this.currentFile = c.file.name;
        var g = this.streamFiles && !c.file.dir;
        if (g) {
          var p = s(c, g, !1, this.currentSourceOffset, this.zipPlatform, this.encodeFileName);
          this.push({ data: p.fileRecord, meta: { percent: 0 } });
        } else this.accumulate = !0;
      }, u.prototype.closedSource = function(c) {
        this.accumulate = !1;
        var g = this.streamFiles && !c.file.dir, p = s(c, g, !0, this.currentSourceOffset, this.zipPlatform, this.encodeFileName);
        if (this.dirRecords.push(p.dirRecord), g) this.push({ data: function(y) {
          return d.DATA_DESCRIPTOR + a(y.crc32, 4) + a(y.compressedSize, 4) + a(y.uncompressedSize, 4);
        }(c), meta: { percent: 100 } });
        else for (this.push({ data: p.fileRecord, meta: { percent: 0 } }); this.contentBuffer.length; ) this.push(this.contentBuffer.shift());
        this.currentFile = null;
      }, u.prototype.flush = function() {
        for (var c = this.bytesWritten, g = 0; g < this.dirRecords.length; g++) this.push({ data: this.dirRecords[g], meta: { percent: 100 } });
        var p = this.bytesWritten - c, y = function(m, A, E, x, S) {
          var R = o.transformTo("string", S(x));
          return d.CENTRAL_DIRECTORY_END + "\0\0\0\0" + a(m, 2) + a(m, 2) + a(A, 4) + a(E, 4) + a(R.length, 2) + R;
        }(this.dirRecords.length, p, c, this.zipComment, this.encodeFileName);
        this.push({ data: y, meta: { percent: 100 } });
      }, u.prototype.prepareNextSource = function() {
        this.previous = this._sources.shift(), this.openedSource(this.previous.streamInfo), this.isPaused ? this.previous.pause() : this.previous.resume();
      }, u.prototype.registerPrevious = function(c) {
        this._sources.push(c);
        var g = this;
        return c.on("data", function(p) {
          g.processChunk(p);
        }), c.on("end", function() {
          g.closedSource(g.previous.streamInfo), g._sources.length ? g.prepareNextSource() : g.end();
        }), c.on("error", function(p) {
          g.error(p);
        }), this;
      }, u.prototype.resume = function() {
        return !!h.prototype.resume.call(this) && (!this.previous && this._sources.length ? (this.prepareNextSource(), !0) : this.previous || this._sources.length || this.generatedError ? void 0 : (this.end(), !0));
      }, u.prototype.error = function(c) {
        var g = this._sources;
        if (!h.prototype.error.call(this, c)) return !1;
        for (var p = 0; p < g.length; p++) try {
          g[p].error(c);
        } catch {
        }
        return !0;
      }, u.prototype.lock = function() {
        h.prototype.lock.call(this);
        for (var c = this._sources, g = 0; g < c.length; g++) c[g].lock();
      }, n.exports = u;
    }, { "../crc32": 4, "../signature": 23, "../stream/GenericWorker": 28, "../utf8": 31, "../utils": 32 }], 9: [function(e, n, i) {
      var a = e("../compressions"), s = e("./ZipFileWorker");
      i.generateWorker = function(o, h, l) {
        var f = new s(h.streamFiles, l, h.platform, h.encodeFileName), d = 0;
        try {
          o.forEach(function(u, c) {
            d++;
            var g = function(A, E) {
              var x = A || E, S = a[x];
              if (!S) throw new Error(x + " is not a valid compression method !");
              return S;
            }(c.options.compression, h.compression), p = c.options.compressionOptions || h.compressionOptions || {}, y = c.dir, m = c.date;
            c._compressWorker(g, p).withStreamInfo("file", { name: u, dir: y, date: m, comment: c.comment || "", unixPermissions: c.unixPermissions, dosPermissions: c.dosPermissions }).pipe(f);
          }), f.entriesCount = d;
        } catch (u) {
          f.error(u);
        }
        return f;
      };
    }, { "../compressions": 3, "./ZipFileWorker": 8 }], 10: [function(e, n, i) {
      function a() {
        if (!(this instanceof a)) return new a();
        if (arguments.length) throw new Error("The constructor with parameters has been removed in JSZip 3.0, please check the upgrade guide.");
        this.files = /* @__PURE__ */ Object.create(null), this.comment = null, this.root = "", this.clone = function() {
          var s = new a();
          for (var o in this) typeof this[o] != "function" && (s[o] = this[o]);
          return s;
        };
      }
      (a.prototype = e("./object")).loadAsync = e("./load"), a.support = e("./support"), a.defaults = e("./defaults"), a.version = "3.10.1", a.loadAsync = function(s, o) {
        return new a().loadAsync(s, o);
      }, a.external = e("./external"), n.exports = a;
    }, { "./defaults": 5, "./external": 6, "./load": 11, "./object": 15, "./support": 30 }], 11: [function(e, n, i) {
      var a = e("./utils"), s = e("./external"), o = e("./utf8"), h = e("./zipEntries"), l = e("./stream/Crc32Probe"), f = e("./nodejsUtils");
      function d(u) {
        return new s.Promise(function(c, g) {
          var p = u.decompressed.getContentWorker().pipe(new l());
          p.on("error", function(y) {
            g(y);
          }).on("end", function() {
            p.streamInfo.crc32 !== u.decompressed.crc32 ? g(new Error("Corrupted zip : CRC32 mismatch")) : c();
          }).resume();
        });
      }
      n.exports = function(u, c) {
        var g = this;
        return c = a.extend(c || {}, { base64: !1, checkCRC32: !1, optimizedBinaryString: !1, createFolders: !1, decodeFileName: o.utf8decode }), f.isNode && f.isStream(u) ? s.Promise.reject(new Error("JSZip can't accept a stream when loading a zip file.")) : a.prepareContent("the loaded zip file", u, !0, c.optimizedBinaryString, c.base64).then(function(p) {
          var y = new h(c);
          return y.load(p), y;
        }).then(function(p) {
          var y = [s.Promise.resolve(p)], m = p.files;
          if (c.checkCRC32) for (var A = 0; A < m.length; A++) y.push(d(m[A]));
          return s.Promise.all(y);
        }).then(function(p) {
          for (var y = p.shift(), m = y.files, A = 0; A < m.length; A++) {
            var E = m[A], x = E.fileNameStr, S = a.resolve(E.fileNameStr);
            g.file(S, E.decompressed, { binary: !0, optimizedBinaryString: !0, date: E.date, dir: E.dir, comment: E.fileCommentStr.length ? E.fileCommentStr : null, unixPermissions: E.unixPermissions, dosPermissions: E.dosPermissions, createFolders: c.createFolders }), E.dir || (g.file(S).unsafeOriginalName = x);
          }
          return y.zipComment.length && (g.comment = y.zipComment), g;
        });
      };
    }, { "./external": 6, "./nodejsUtils": 14, "./stream/Crc32Probe": 25, "./utf8": 31, "./utils": 32, "./zipEntries": 33 }], 12: [function(e, n, i) {
      var a = e("../utils"), s = e("../stream/GenericWorker");
      function o(h, l) {
        s.call(this, "Nodejs stream input adapter for " + h), this._upstreamEnded = !1, this._bindStream(l);
      }
      a.inherits(o, s), o.prototype._bindStream = function(h) {
        var l = this;
        (this._stream = h).pause(), h.on("data", function(f) {
          l.push({ data: f, meta: { percent: 0 } });
        }).on("error", function(f) {
          l.isPaused ? this.generatedError = f : l.error(f);
        }).on("end", function() {
          l.isPaused ? l._upstreamEnded = !0 : l.end();
        });
      }, o.prototype.pause = function() {
        return !!s.prototype.pause.call(this) && (this._stream.pause(), !0);
      }, o.prototype.resume = function() {
        return !!s.prototype.resume.call(this) && (this._upstreamEnded ? this.end() : this._stream.resume(), !0);
      }, n.exports = o;
    }, { "../stream/GenericWorker": 28, "../utils": 32 }], 13: [function(e, n, i) {
      var a = e("readable-stream").Readable;
      function s(o, h, l) {
        a.call(this, h), this._helper = o;
        var f = this;
        o.on("data", function(d, u) {
          f.push(d) || f._helper.pause(), l && l(u);
        }).on("error", function(d) {
          f.emit("error", d);
        }).on("end", function() {
          f.push(null);
        });
      }
      e("../utils").inherits(s, a), s.prototype._read = function() {
        this._helper.resume();
      }, n.exports = s;
    }, { "../utils": 32, "readable-stream": 16 }], 14: [function(e, n, i) {
      n.exports = { isNode: typeof Buffer < "u", newBufferFrom: function(a, s) {
        if (Buffer.from && Buffer.from !== Uint8Array.from) return Buffer.from(a, s);
        if (typeof a == "number") throw new Error('The "data" argument must not be a number');
        return new Buffer(a, s);
      }, allocBuffer: function(a) {
        if (Buffer.alloc) return Buffer.alloc(a);
        var s = new Buffer(a);
        return s.fill(0), s;
      }, isBuffer: function(a) {
        return Buffer.isBuffer(a);
      }, isStream: function(a) {
        return a && typeof a.on == "function" && typeof a.pause == "function" && typeof a.resume == "function";
      } };
    }, {}], 15: [function(e, n, i) {
      function a(S, R, v) {
        var O, C = o.getTypeOf(R), B = o.extend(v || {}, f);
        B.date = B.date || /* @__PURE__ */ new Date(), B.compression !== null && (B.compression = B.compression.toUpperCase()), typeof B.unixPermissions == "string" && (B.unixPermissions = parseInt(B.unixPermissions, 8)), B.unixPermissions && 16384 & B.unixPermissions && (B.dir = !0), B.dosPermissions && 16 & B.dosPermissions && (B.dir = !0), B.dir && (S = m(S)), B.createFolders && (O = y(S)) && A.call(this, O, !0);
        var Z = C === "string" && B.binary === !1 && B.base64 === !1;
        v && v.binary !== void 0 || (B.binary = !Z), (R instanceof d && R.uncompressedSize === 0 || B.dir || !R || R.length === 0) && (B.base64 = !1, B.binary = !0, R = "", B.compression = "STORE", C = "string");
        var T = null;
        T = R instanceof d || R instanceof h ? R : g.isNode && g.isStream(R) ? new p(S, R) : o.prepareContent(S, R, B.binary, B.optimizedBinaryString, B.base64);
        var j = new u(S, T, B);
        this.files[S] = j;
      }
      var s = e("./utf8"), o = e("./utils"), h = e("./stream/GenericWorker"), l = e("./stream/StreamHelper"), f = e("./defaults"), d = e("./compressedObject"), u = e("./zipObject"), c = e("./generate"), g = e("./nodejsUtils"), p = e("./nodejs/NodejsStreamInputAdapter"), y = function(S) {
        S.slice(-1) === "/" && (S = S.substring(0, S.length - 1));
        var R = S.lastIndexOf("/");
        return 0 < R ? S.substring(0, R) : "";
      }, m = function(S) {
        return S.slice(-1) !== "/" && (S += "/"), S;
      }, A = function(S, R) {
        return R = R !== void 0 ? R : f.createFolders, S = m(S), this.files[S] || a.call(this, S, null, { dir: !0, createFolders: R }), this.files[S];
      };
      function E(S) {
        return Object.prototype.toString.call(S) === "[object RegExp]";
      }
      var x = { load: function() {
        throw new Error("This method has been removed in JSZip 3.0, please check the upgrade guide.");
      }, forEach: function(S) {
        var R, v, O;
        for (R in this.files) O = this.files[R], (v = R.slice(this.root.length, R.length)) && R.slice(0, this.root.length) === this.root && S(v, O);
      }, filter: function(S) {
        var R = [];
        return this.forEach(function(v, O) {
          S(v, O) && R.push(O);
        }), R;
      }, file: function(S, R, v) {
        if (arguments.length !== 1) return S = this.root + S, a.call(this, S, R, v), this;
        if (E(S)) {
          var O = S;
          return this.filter(function(B, Z) {
            return !Z.dir && O.test(B);
          });
        }
        var C = this.files[this.root + S];
        return C && !C.dir ? C : null;
      }, folder: function(S) {
        if (!S) return this;
        if (E(S)) return this.filter(function(C, B) {
          return B.dir && S.test(C);
        });
        var R = this.root + S, v = A.call(this, R), O = this.clone();
        return O.root = v.name, O;
      }, remove: function(S) {
        S = this.root + S;
        var R = this.files[S];
        if (R || (S.slice(-1) !== "/" && (S += "/"), R = this.files[S]), R && !R.dir) delete this.files[S];
        else for (var v = this.filter(function(C, B) {
          return B.name.slice(0, S.length) === S;
        }), O = 0; O < v.length; O++) delete this.files[v[O].name];
        return this;
      }, generate: function() {
        throw new Error("This method has been removed in JSZip 3.0, please check the upgrade guide.");
      }, generateInternalStream: function(S) {
        var R, v = {};
        try {
          if ((v = o.extend(S || {}, { streamFiles: !1, compression: "STORE", compressionOptions: null, type: "", platform: "DOS", comment: null, mimeType: "application/zip", encodeFileName: s.utf8encode })).type = v.type.toLowerCase(), v.compression = v.compression.toUpperCase(), v.type === "binarystring" && (v.type = "string"), !v.type) throw new Error("No output type specified.");
          o.checkSupport(v.type), v.platform !== "darwin" && v.platform !== "freebsd" && v.platform !== "linux" && v.platform !== "sunos" || (v.platform = "UNIX"), v.platform === "win32" && (v.platform = "DOS");
          var O = v.comment || this.comment || "";
          R = c.generateWorker(this, v, O);
        } catch (C) {
          (R = new h("error")).error(C);
        }
        return new l(R, v.type || "string", v.mimeType);
      }, generateAsync: function(S, R) {
        return this.generateInternalStream(S).accumulate(R);
      }, generateNodeStream: function(S, R) {
        return (S = S || {}).type || (S.type = "nodebuffer"), this.generateInternalStream(S).toNodejsStream(R);
      } };
      n.exports = x;
    }, { "./compressedObject": 2, "./defaults": 5, "./generate": 9, "./nodejs/NodejsStreamInputAdapter": 12, "./nodejsUtils": 14, "./stream/GenericWorker": 28, "./stream/StreamHelper": 29, "./utf8": 31, "./utils": 32, "./zipObject": 35 }], 16: [function(e, n, i) {
      n.exports = e("stream");
    }, { stream: void 0 }], 17: [function(e, n, i) {
      var a = e("./DataReader");
      function s(o) {
        a.call(this, o);
        for (var h = 0; h < this.data.length; h++) o[h] = 255 & o[h];
      }
      e("../utils").inherits(s, a), s.prototype.byteAt = function(o) {
        return this.data[this.zero + o];
      }, s.prototype.lastIndexOfSignature = function(o) {
        for (var h = o.charCodeAt(0), l = o.charCodeAt(1), f = o.charCodeAt(2), d = o.charCodeAt(3), u = this.length - 4; 0 <= u; --u) if (this.data[u] === h && this.data[u + 1] === l && this.data[u + 2] === f && this.data[u + 3] === d) return u - this.zero;
        return -1;
      }, s.prototype.readAndCheckSignature = function(o) {
        var h = o.charCodeAt(0), l = o.charCodeAt(1), f = o.charCodeAt(2), d = o.charCodeAt(3), u = this.readData(4);
        return h === u[0] && l === u[1] && f === u[2] && d === u[3];
      }, s.prototype.readData = function(o) {
        if (this.checkOffset(o), o === 0) return [];
        var h = this.data.slice(this.zero + this.index, this.zero + this.index + o);
        return this.index += o, h;
      }, n.exports = s;
    }, { "../utils": 32, "./DataReader": 18 }], 18: [function(e, n, i) {
      var a = e("../utils");
      function s(o) {
        this.data = o, this.length = o.length, this.index = 0, this.zero = 0;
      }
      s.prototype = { checkOffset: function(o) {
        this.checkIndex(this.index + o);
      }, checkIndex: function(o) {
        if (this.length < this.zero + o || o < 0) throw new Error("End of data reached (data length = " + this.length + ", asked index = " + o + "). Corrupted zip ?");
      }, setIndex: function(o) {
        this.checkIndex(o), this.index = o;
      }, skip: function(o) {
        this.setIndex(this.index + o);
      }, byteAt: function() {
      }, readInt: function(o) {
        var h, l = 0;
        for (this.checkOffset(o), h = this.index + o - 1; h >= this.index; h--) l = (l << 8) + this.byteAt(h);
        return this.index += o, l;
      }, readString: function(o) {
        return a.transformTo("string", this.readData(o));
      }, readData: function() {
      }, lastIndexOfSignature: function() {
      }, readAndCheckSignature: function() {
      }, readDate: function() {
        var o = this.readInt(4);
        return new Date(Date.UTC(1980 + (o >> 25 & 127), (o >> 21 & 15) - 1, o >> 16 & 31, o >> 11 & 31, o >> 5 & 63, (31 & o) << 1));
      } }, n.exports = s;
    }, { "../utils": 32 }], 19: [function(e, n, i) {
      var a = e("./Uint8ArrayReader");
      function s(o) {
        a.call(this, o);
      }
      e("../utils").inherits(s, a), s.prototype.readData = function(o) {
        this.checkOffset(o);
        var h = this.data.slice(this.zero + this.index, this.zero + this.index + o);
        return this.index += o, h;
      }, n.exports = s;
    }, { "../utils": 32, "./Uint8ArrayReader": 21 }], 20: [function(e, n, i) {
      var a = e("./DataReader");
      function s(o) {
        a.call(this, o);
      }
      e("../utils").inherits(s, a), s.prototype.byteAt = function(o) {
        return this.data.charCodeAt(this.zero + o);
      }, s.prototype.lastIndexOfSignature = function(o) {
        return this.data.lastIndexOf(o) - this.zero;
      }, s.prototype.readAndCheckSignature = function(o) {
        return o === this.readData(4);
      }, s.prototype.readData = function(o) {
        this.checkOffset(o);
        var h = this.data.slice(this.zero + this.index, this.zero + this.index + o);
        return this.index += o, h;
      }, n.exports = s;
    }, { "../utils": 32, "./DataReader": 18 }], 21: [function(e, n, i) {
      var a = e("./ArrayReader");
      function s(o) {
        a.call(this, o);
      }
      e("../utils").inherits(s, a), s.prototype.readData = function(o) {
        if (this.checkOffset(o), o === 0) return new Uint8Array(0);
        var h = this.data.subarray(this.zero + this.index, this.zero + this.index + o);
        return this.index += o, h;
      }, n.exports = s;
    }, { "../utils": 32, "./ArrayReader": 17 }], 22: [function(e, n, i) {
      var a = e("../utils"), s = e("../support"), o = e("./ArrayReader"), h = e("./StringReader"), l = e("./NodeBufferReader"), f = e("./Uint8ArrayReader");
      n.exports = function(d) {
        var u = a.getTypeOf(d);
        return a.checkSupport(u), u !== "string" || s.uint8array ? u === "nodebuffer" ? new l(d) : s.uint8array ? new f(a.transformTo("uint8array", d)) : new o(a.transformTo("array", d)) : new h(d);
      };
    }, { "../support": 30, "../utils": 32, "./ArrayReader": 17, "./NodeBufferReader": 19, "./StringReader": 20, "./Uint8ArrayReader": 21 }], 23: [function(e, n, i) {
      i.LOCAL_FILE_HEADER = "PK", i.CENTRAL_FILE_HEADER = "PK", i.CENTRAL_DIRECTORY_END = "PK", i.ZIP64_CENTRAL_DIRECTORY_LOCATOR = "PK\x07", i.ZIP64_CENTRAL_DIRECTORY_END = "PK", i.DATA_DESCRIPTOR = "PK\x07\b";
    }, {}], 24: [function(e, n, i) {
      var a = e("./GenericWorker"), s = e("../utils");
      function o(h) {
        a.call(this, "ConvertWorker to " + h), this.destType = h;
      }
      s.inherits(o, a), o.prototype.processChunk = function(h) {
        this.push({ data: s.transformTo(this.destType, h.data), meta: h.meta });
      }, n.exports = o;
    }, { "../utils": 32, "./GenericWorker": 28 }], 25: [function(e, n, i) {
      var a = e("./GenericWorker"), s = e("../crc32");
      function o() {
        a.call(this, "Crc32Probe"), this.withStreamInfo("crc32", 0);
      }
      e("../utils").inherits(o, a), o.prototype.processChunk = function(h) {
        this.streamInfo.crc32 = s(h.data, this.streamInfo.crc32 || 0), this.push(h);
      }, n.exports = o;
    }, { "../crc32": 4, "../utils": 32, "./GenericWorker": 28 }], 26: [function(e, n, i) {
      var a = e("../utils"), s = e("./GenericWorker");
      function o(h) {
        s.call(this, "DataLengthProbe for " + h), this.propName = h, this.withStreamInfo(h, 0);
      }
      a.inherits(o, s), o.prototype.processChunk = function(h) {
        if (h) {
          var l = this.streamInfo[this.propName] || 0;
          this.streamInfo[this.propName] = l + h.data.length;
        }
        s.prototype.processChunk.call(this, h);
      }, n.exports = o;
    }, { "../utils": 32, "./GenericWorker": 28 }], 27: [function(e, n, i) {
      var a = e("../utils"), s = e("./GenericWorker");
      function o(h) {
        s.call(this, "DataWorker");
        var l = this;
        this.dataIsReady = !1, this.index = 0, this.max = 0, this.data = null, this.type = "", this._tickScheduled = !1, h.then(function(f) {
          l.dataIsReady = !0, l.data = f, l.max = f && f.length || 0, l.type = a.getTypeOf(f), l.isPaused || l._tickAndRepeat();
        }, function(f) {
          l.error(f);
        });
      }
      a.inherits(o, s), o.prototype.cleanUp = function() {
        s.prototype.cleanUp.call(this), this.data = null;
      }, o.prototype.resume = function() {
        return !!s.prototype.resume.call(this) && (!this._tickScheduled && this.dataIsReady && (this._tickScheduled = !0, a.delay(this._tickAndRepeat, [], this)), !0);
      }, o.prototype._tickAndRepeat = function() {
        this._tickScheduled = !1, this.isPaused || this.isFinished || (this._tick(), this.isFinished || (a.delay(this._tickAndRepeat, [], this), this._tickScheduled = !0));
      }, o.prototype._tick = function() {
        if (this.isPaused || this.isFinished) return !1;
        var h = null, l = Math.min(this.max, this.index + 16384);
        if (this.index >= this.max) return this.end();
        switch (this.type) {
          case "string":
            h = this.data.substring(this.index, l);
            break;
          case "uint8array":
            h = this.data.subarray(this.index, l);
            break;
          case "array":
          case "nodebuffer":
            h = this.data.slice(this.index, l);
        }
        return this.index = l, this.push({ data: h, meta: { percent: this.max ? this.index / this.max * 100 : 0 } });
      }, n.exports = o;
    }, { "../utils": 32, "./GenericWorker": 28 }], 28: [function(e, n, i) {
      function a(s) {
        this.name = s || "default", this.streamInfo = {}, this.generatedError = null, this.extraStreamInfo = {}, this.isPaused = !0, this.isFinished = !1, this.isLocked = !1, this._listeners = { data: [], end: [], error: [] }, this.previous = null;
      }
      a.prototype = { push: function(s) {
        this.emit("data", s);
      }, end: function() {
        if (this.isFinished) return !1;
        this.flush();
        try {
          this.emit("end"), this.cleanUp(), this.isFinished = !0;
        } catch (s) {
          this.emit("error", s);
        }
        return !0;
      }, error: function(s) {
        return !this.isFinished && (this.isPaused ? this.generatedError = s : (this.isFinished = !0, this.emit("error", s), this.previous && this.previous.error(s), this.cleanUp()), !0);
      }, on: function(s, o) {
        return this._listeners[s].push(o), this;
      }, cleanUp: function() {
        this.streamInfo = this.generatedError = this.extraStreamInfo = null, this._listeners = [];
      }, emit: function(s, o) {
        if (this._listeners[s]) for (var h = 0; h < this._listeners[s].length; h++) this._listeners[s][h].call(this, o);
      }, pipe: function(s) {
        return s.registerPrevious(this);
      }, registerPrevious: function(s) {
        if (this.isLocked) throw new Error("The stream '" + this + "' has already been used.");
        this.streamInfo = s.streamInfo, this.mergeStreamInfo(), this.previous = s;
        var o = this;
        return s.on("data", function(h) {
          o.processChunk(h);
        }), s.on("end", function() {
          o.end();
        }), s.on("error", function(h) {
          o.error(h);
        }), this;
      }, pause: function() {
        return !this.isPaused && !this.isFinished && (this.isPaused = !0, this.previous && this.previous.pause(), !0);
      }, resume: function() {
        if (!this.isPaused || this.isFinished) return !1;
        var s = this.isPaused = !1;
        return this.generatedError && (this.error(this.generatedError), s = !0), this.previous && this.previous.resume(), !s;
      }, flush: function() {
      }, processChunk: function(s) {
        this.push(s);
      }, withStreamInfo: function(s, o) {
        return this.extraStreamInfo[s] = o, this.mergeStreamInfo(), this;
      }, mergeStreamInfo: function() {
        for (var s in this.extraStreamInfo) Object.prototype.hasOwnProperty.call(this.extraStreamInfo, s) && (this.streamInfo[s] = this.extraStreamInfo[s]);
      }, lock: function() {
        if (this.isLocked) throw new Error("The stream '" + this + "' has already been used.");
        this.isLocked = !0, this.previous && this.previous.lock();
      }, toString: function() {
        var s = "Worker " + this.name;
        return this.previous ? this.previous + " -> " + s : s;
      } }, n.exports = a;
    }, {}], 29: [function(e, n, i) {
      var a = e("../utils"), s = e("./ConvertWorker"), o = e("./GenericWorker"), h = e("../base64"), l = e("../support"), f = e("../external"), d = null;
      if (l.nodestream) try {
        d = e("../nodejs/NodejsStreamOutputAdapter");
      } catch {
      }
      function u(g, p) {
        return new f.Promise(function(y, m) {
          var A = [], E = g._internalType, x = g._outputType, S = g._mimeType;
          g.on("data", function(R, v) {
            A.push(R), p && p(v);
          }).on("error", function(R) {
            A = [], m(R);
          }).on("end", function() {
            try {
              var R = function(v, O, C) {
                switch (v) {
                  case "blob":
                    return a.newBlob(a.transformTo("arraybuffer", O), C);
                  case "base64":
                    return h.encode(O);
                  default:
                    return a.transformTo(v, O);
                }
              }(x, function(v, O) {
                var C, B = 0, Z = null, T = 0;
                for (C = 0; C < O.length; C++) T += O[C].length;
                switch (v) {
                  case "string":
                    return O.join("");
                  case "array":
                    return Array.prototype.concat.apply([], O);
                  case "uint8array":
                    for (Z = new Uint8Array(T), C = 0; C < O.length; C++) Z.set(O[C], B), B += O[C].length;
                    return Z;
                  case "nodebuffer":
                    return Buffer.concat(O);
                  default:
                    throw new Error("concat : unsupported type '" + v + "'");
                }
              }(E, A), S);
              y(R);
            } catch (v) {
              m(v);
            }
            A = [];
          }).resume();
        });
      }
      function c(g, p, y) {
        var m = p;
        switch (p) {
          case "blob":
          case "arraybuffer":
            m = "uint8array";
            break;
          case "base64":
            m = "string";
        }
        try {
          this._internalType = m, this._outputType = p, this._mimeType = y, a.checkSupport(m), this._worker = g.pipe(new s(m)), g.lock();
        } catch (A) {
          this._worker = new o("error"), this._worker.error(A);
        }
      }
      c.prototype = { accumulate: function(g) {
        return u(this, g);
      }, on: function(g, p) {
        var y = this;
        return g === "data" ? this._worker.on(g, function(m) {
          p.call(y, m.data, m.meta);
        }) : this._worker.on(g, function() {
          a.delay(p, arguments, y);
        }), this;
      }, resume: function() {
        return a.delay(this._worker.resume, [], this._worker), this;
      }, pause: function() {
        return this._worker.pause(), this;
      }, toNodejsStream: function(g) {
        if (a.checkSupport("nodestream"), this._outputType !== "nodebuffer") throw new Error(this._outputType + " is not supported by this method");
        return new d(this, { objectMode: this._outputType !== "nodebuffer" }, g);
      } }, n.exports = c;
    }, { "../base64": 1, "../external": 6, "../nodejs/NodejsStreamOutputAdapter": 13, "../support": 30, "../utils": 32, "./ConvertWorker": 24, "./GenericWorker": 28 }], 30: [function(e, n, i) {
      if (i.base64 = !0, i.array = !0, i.string = !0, i.arraybuffer = typeof ArrayBuffer < "u" && typeof Uint8Array < "u", i.nodebuffer = typeof Buffer < "u", i.uint8array = typeof Uint8Array < "u", typeof ArrayBuffer > "u") i.blob = !1;
      else {
        var a = new ArrayBuffer(0);
        try {
          i.blob = new Blob([a], { type: "application/zip" }).size === 0;
        } catch {
          try {
            var s = new (self.BlobBuilder || self.WebKitBlobBuilder || self.MozBlobBuilder || self.MSBlobBuilder)();
            s.append(a), i.blob = s.getBlob("application/zip").size === 0;
          } catch {
            i.blob = !1;
          }
        }
      }
      try {
        i.nodestream = !!e("readable-stream").Readable;
      } catch {
        i.nodestream = !1;
      }
    }, { "readable-stream": 16 }], 31: [function(e, n, i) {
      for (var a = e("./utils"), s = e("./support"), o = e("./nodejsUtils"), h = e("./stream/GenericWorker"), l = new Array(256), f = 0; f < 256; f++) l[f] = 252 <= f ? 6 : 248 <= f ? 5 : 240 <= f ? 4 : 224 <= f ? 3 : 192 <= f ? 2 : 1;
      l[254] = l[254] = 1;
      function d() {
        h.call(this, "utf-8 decode"), this.leftOver = null;
      }
      function u() {
        h.call(this, "utf-8 encode");
      }
      i.utf8encode = function(c) {
        return s.nodebuffer ? o.newBufferFrom(c, "utf-8") : function(g) {
          var p, y, m, A, E, x = g.length, S = 0;
          for (A = 0; A < x; A++) (64512 & (y = g.charCodeAt(A))) == 55296 && A + 1 < x && (64512 & (m = g.charCodeAt(A + 1))) == 56320 && (y = 65536 + (y - 55296 << 10) + (m - 56320), A++), S += y < 128 ? 1 : y < 2048 ? 2 : y < 65536 ? 3 : 4;
          for (p = s.uint8array ? new Uint8Array(S) : new Array(S), A = E = 0; E < S; A++) (64512 & (y = g.charCodeAt(A))) == 55296 && A + 1 < x && (64512 & (m = g.charCodeAt(A + 1))) == 56320 && (y = 65536 + (y - 55296 << 10) + (m - 56320), A++), y < 128 ? p[E++] = y : (y < 2048 ? p[E++] = 192 | y >>> 6 : (y < 65536 ? p[E++] = 224 | y >>> 12 : (p[E++] = 240 | y >>> 18, p[E++] = 128 | y >>> 12 & 63), p[E++] = 128 | y >>> 6 & 63), p[E++] = 128 | 63 & y);
          return p;
        }(c);
      }, i.utf8decode = function(c) {
        return s.nodebuffer ? a.transformTo("nodebuffer", c).toString("utf-8") : function(g) {
          var p, y, m, A, E = g.length, x = new Array(2 * E);
          for (p = y = 0; p < E; ) if ((m = g[p++]) < 128) x[y++] = m;
          else if (4 < (A = l[m])) x[y++] = 65533, p += A - 1;
          else {
            for (m &= A === 2 ? 31 : A === 3 ? 15 : 7; 1 < A && p < E; ) m = m << 6 | 63 & g[p++], A--;
            1 < A ? x[y++] = 65533 : m < 65536 ? x[y++] = m : (m -= 65536, x[y++] = 55296 | m >> 10 & 1023, x[y++] = 56320 | 1023 & m);
          }
          return x.length !== y && (x.subarray ? x = x.subarray(0, y) : x.length = y), a.applyFromCharCode(x);
        }(c = a.transformTo(s.uint8array ? "uint8array" : "array", c));
      }, a.inherits(d, h), d.prototype.processChunk = function(c) {
        var g = a.transformTo(s.uint8array ? "uint8array" : "array", c.data);
        if (this.leftOver && this.leftOver.length) {
          if (s.uint8array) {
            var p = g;
            (g = new Uint8Array(p.length + this.leftOver.length)).set(this.leftOver, 0), g.set(p, this.leftOver.length);
          } else g = this.leftOver.concat(g);
          this.leftOver = null;
        }
        var y = function(A, E) {
          var x;
          for ((E = E || A.length) > A.length && (E = A.length), x = E - 1; 0 <= x && (192 & A[x]) == 128; ) x--;
          return x < 0 || x === 0 ? E : x + l[A[x]] > E ? x : E;
        }(g), m = g;
        y !== g.length && (s.uint8array ? (m = g.subarray(0, y), this.leftOver = g.subarray(y, g.length)) : (m = g.slice(0, y), this.leftOver = g.slice(y, g.length))), this.push({ data: i.utf8decode(m), meta: c.meta });
      }, d.prototype.flush = function() {
        this.leftOver && this.leftOver.length && (this.push({ data: i.utf8decode(this.leftOver), meta: {} }), this.leftOver = null);
      }, i.Utf8DecodeWorker = d, a.inherits(u, h), u.prototype.processChunk = function(c) {
        this.push({ data: i.utf8encode(c.data), meta: c.meta });
      }, i.Utf8EncodeWorker = u;
    }, { "./nodejsUtils": 14, "./stream/GenericWorker": 28, "./support": 30, "./utils": 32 }], 32: [function(e, n, i) {
      var a = e("./support"), s = e("./base64"), o = e("./nodejsUtils"), h = e("./external");
      function l(p) {
        return p;
      }
      function f(p, y) {
        for (var m = 0; m < p.length; ++m) y[m] = 255 & p.charCodeAt(m);
        return y;
      }
      e("setimmediate"), i.newBlob = function(p, y) {
        i.checkSupport("blob");
        try {
          return new Blob([p], { type: y });
        } catch {
          try {
            var m = new (self.BlobBuilder || self.WebKitBlobBuilder || self.MozBlobBuilder || self.MSBlobBuilder)();
            return m.append(p), m.getBlob(y);
          } catch {
            throw new Error("Bug : can't construct the Blob.");
          }
        }
      };
      var d = { stringifyByChunk: function(p, y, m) {
        var A = [], E = 0, x = p.length;
        if (x <= m) return String.fromCharCode.apply(null, p);
        for (; E < x; ) y === "array" || y === "nodebuffer" ? A.push(String.fromCharCode.apply(null, p.slice(E, Math.min(E + m, x)))) : A.push(String.fromCharCode.apply(null, p.subarray(E, Math.min(E + m, x)))), E += m;
        return A.join("");
      }, stringifyByChar: function(p) {
        for (var y = "", m = 0; m < p.length; m++) y += String.fromCharCode(p[m]);
        return y;
      }, applyCanBeUsed: { uint8array: function() {
        try {
          return a.uint8array && String.fromCharCode.apply(null, new Uint8Array(1)).length === 1;
        } catch {
          return !1;
        }
      }(), nodebuffer: function() {
        try {
          return a.nodebuffer && String.fromCharCode.apply(null, o.allocBuffer(1)).length === 1;
        } catch {
          return !1;
        }
      }() } };
      function u(p) {
        var y = 65536, m = i.getTypeOf(p), A = !0;
        if (m === "uint8array" ? A = d.applyCanBeUsed.uint8array : m === "nodebuffer" && (A = d.applyCanBeUsed.nodebuffer), A) for (; 1 < y; ) try {
          return d.stringifyByChunk(p, m, y);
        } catch {
          y = Math.floor(y / 2);
        }
        return d.stringifyByChar(p);
      }
      function c(p, y) {
        for (var m = 0; m < p.length; m++) y[m] = p[m];
        return y;
      }
      i.applyFromCharCode = u;
      var g = {};
      g.string = { string: l, array: function(p) {
        return f(p, new Array(p.length));
      }, arraybuffer: function(p) {
        return g.string.uint8array(p).buffer;
      }, uint8array: function(p) {
        return f(p, new Uint8Array(p.length));
      }, nodebuffer: function(p) {
        return f(p, o.allocBuffer(p.length));
      } }, g.array = { string: u, array: l, arraybuffer: function(p) {
        return new Uint8Array(p).buffer;
      }, uint8array: function(p) {
        return new Uint8Array(p);
      }, nodebuffer: function(p) {
        return o.newBufferFrom(p);
      } }, g.arraybuffer = { string: function(p) {
        return u(new Uint8Array(p));
      }, array: function(p) {
        return c(new Uint8Array(p), new Array(p.byteLength));
      }, arraybuffer: l, uint8array: function(p) {
        return new Uint8Array(p);
      }, nodebuffer: function(p) {
        return o.newBufferFrom(new Uint8Array(p));
      } }, g.uint8array = { string: u, array: function(p) {
        return c(p, new Array(p.length));
      }, arraybuffer: function(p) {
        return p.buffer;
      }, uint8array: l, nodebuffer: function(p) {
        return o.newBufferFrom(p);
      } }, g.nodebuffer = { string: u, array: function(p) {
        return c(p, new Array(p.length));
      }, arraybuffer: function(p) {
        return g.nodebuffer.uint8array(p).buffer;
      }, uint8array: function(p) {
        return c(p, new Uint8Array(p.length));
      }, nodebuffer: l }, i.transformTo = function(p, y) {
        if (y = y || "", !p) return y;
        i.checkSupport(p);
        var m = i.getTypeOf(y);
        return g[m][p](y);
      }, i.resolve = function(p) {
        for (var y = p.split("/"), m = [], A = 0; A < y.length; A++) {
          var E = y[A];
          E === "." || E === "" && A !== 0 && A !== y.length - 1 || (E === ".." ? m.pop() : m.push(E));
        }
        return m.join("/");
      }, i.getTypeOf = function(p) {
        return typeof p == "string" ? "string" : Object.prototype.toString.call(p) === "[object Array]" ? "array" : a.nodebuffer && o.isBuffer(p) ? "nodebuffer" : a.uint8array && p instanceof Uint8Array ? "uint8array" : a.arraybuffer && p instanceof ArrayBuffer ? "arraybuffer" : void 0;
      }, i.checkSupport = function(p) {
        if (!a[p.toLowerCase()]) throw new Error(p + " is not supported by this platform");
      }, i.MAX_VALUE_16BITS = 65535, i.MAX_VALUE_32BITS = -1, i.pretty = function(p) {
        var y, m, A = "";
        for (m = 0; m < (p || "").length; m++) A += "\\x" + ((y = p.charCodeAt(m)) < 16 ? "0" : "") + y.toString(16).toUpperCase();
        return A;
      }, i.delay = function(p, y, m) {
        setImmediate(function() {
          p.apply(m || null, y || []);
        });
      }, i.inherits = function(p, y) {
        function m() {
        }
        m.prototype = y.prototype, p.prototype = new m();
      }, i.extend = function() {
        var p, y, m = {};
        for (p = 0; p < arguments.length; p++) for (y in arguments[p]) Object.prototype.hasOwnProperty.call(arguments[p], y) && m[y] === void 0 && (m[y] = arguments[p][y]);
        return m;
      }, i.prepareContent = function(p, y, m, A, E) {
        return h.Promise.resolve(y).then(function(x) {
          return a.blob && (x instanceof Blob || ["[object File]", "[object Blob]"].indexOf(Object.prototype.toString.call(x)) !== -1) && typeof FileReader < "u" ? new h.Promise(function(S, R) {
            var v = new FileReader();
            v.onload = function(O) {
              S(O.target.result);
            }, v.onerror = function(O) {
              R(O.target.error);
            }, v.readAsArrayBuffer(x);
          }) : x;
        }).then(function(x) {
          var S = i.getTypeOf(x);
          return S ? (S === "arraybuffer" ? x = i.transformTo("uint8array", x) : S === "string" && (E ? x = s.decode(x) : m && A !== !0 && (x = function(R) {
            return f(R, a.uint8array ? new Uint8Array(R.length) : new Array(R.length));
          }(x))), x) : h.Promise.reject(new Error("Can't read the data of '" + p + "'. Is it in a supported JavaScript type (String, Blob, ArrayBuffer, etc) ?"));
        });
      };
    }, { "./base64": 1, "./external": 6, "./nodejsUtils": 14, "./support": 30, setimmediate: 54 }], 33: [function(e, n, i) {
      var a = e("./reader/readerFor"), s = e("./utils"), o = e("./signature"), h = e("./zipEntry"), l = e("./support");
      function f(d) {
        this.files = [], this.loadOptions = d;
      }
      f.prototype = { checkSignature: function(d) {
        if (!this.reader.readAndCheckSignature(d)) {
          this.reader.index -= 4;
          var u = this.reader.readString(4);
          throw new Error("Corrupted zip or bug: unexpected signature (" + s.pretty(u) + ", expected " + s.pretty(d) + ")");
        }
      }, isSignature: function(d, u) {
        var c = this.reader.index;
        this.reader.setIndex(d);
        var g = this.reader.readString(4) === u;
        return this.reader.setIndex(c), g;
      }, readBlockEndOfCentral: function() {
        this.diskNumber = this.reader.readInt(2), this.diskWithCentralDirStart = this.reader.readInt(2), this.centralDirRecordsOnThisDisk = this.reader.readInt(2), this.centralDirRecords = this.reader.readInt(2), this.centralDirSize = this.reader.readInt(4), this.centralDirOffset = this.reader.readInt(4), this.zipCommentLength = this.reader.readInt(2);
        var d = this.reader.readData(this.zipCommentLength), u = l.uint8array ? "uint8array" : "array", c = s.transformTo(u, d);
        this.zipComment = this.loadOptions.decodeFileName(c);
      }, readBlockZip64EndOfCentral: function() {
        this.zip64EndOfCentralSize = this.reader.readInt(8), this.reader.skip(4), this.diskNumber = this.reader.readInt(4), this.diskWithCentralDirStart = this.reader.readInt(4), this.centralDirRecordsOnThisDisk = this.reader.readInt(8), this.centralDirRecords = this.reader.readInt(8), this.centralDirSize = this.reader.readInt(8), this.centralDirOffset = this.reader.readInt(8), this.zip64ExtensibleData = {};
        for (var d, u, c, g = this.zip64EndOfCentralSize - 44; 0 < g; ) d = this.reader.readInt(2), u = this.reader.readInt(4), c = this.reader.readData(u), this.zip64ExtensibleData[d] = { id: d, length: u, value: c };
      }, readBlockZip64EndOfCentralLocator: function() {
        if (this.diskWithZip64CentralDirStart = this.reader.readInt(4), this.relativeOffsetEndOfZip64CentralDir = this.reader.readInt(8), this.disksCount = this.reader.readInt(4), 1 < this.disksCount) throw new Error("Multi-volumes zip are not supported");
      }, readLocalFiles: function() {
        var d, u;
        for (d = 0; d < this.files.length; d++) u = this.files[d], this.reader.setIndex(u.localHeaderOffset), this.checkSignature(o.LOCAL_FILE_HEADER), u.readLocalPart(this.reader), u.handleUTF8(), u.processAttributes();
      }, readCentralDir: function() {
        var d;
        for (this.reader.setIndex(this.centralDirOffset); this.reader.readAndCheckSignature(o.CENTRAL_FILE_HEADER); ) (d = new h({ zip64: this.zip64 }, this.loadOptions)).readCentralPart(this.reader), this.files.push(d);
        if (this.centralDirRecords !== this.files.length && this.centralDirRecords !== 0 && this.files.length === 0) throw new Error("Corrupted zip or bug: expected " + this.centralDirRecords + " records in central dir, got " + this.files.length);
      }, readEndOfCentral: function() {
        var d = this.reader.lastIndexOfSignature(o.CENTRAL_DIRECTORY_END);
        if (d < 0) throw this.isSignature(0, o.LOCAL_FILE_HEADER) ? new Error("Corrupted zip: can't find end of central directory") : new Error("Can't find end of central directory : is this a zip file ? If it is, see https://stuk.github.io/jszip/documentation/howto/read_zip.html");
        this.reader.setIndex(d);
        var u = d;
        if (this.checkSignature(o.CENTRAL_DIRECTORY_END), this.readBlockEndOfCentral(), this.diskNumber === s.MAX_VALUE_16BITS || this.diskWithCentralDirStart === s.MAX_VALUE_16BITS || this.centralDirRecordsOnThisDisk === s.MAX_VALUE_16BITS || this.centralDirRecords === s.MAX_VALUE_16BITS || this.centralDirSize === s.MAX_VALUE_32BITS || this.centralDirOffset === s.MAX_VALUE_32BITS) {
          if (this.zip64 = !0, (d = this.reader.lastIndexOfSignature(o.ZIP64_CENTRAL_DIRECTORY_LOCATOR)) < 0) throw new Error("Corrupted zip: can't find the ZIP64 end of central directory locator");
          if (this.reader.setIndex(d), this.checkSignature(o.ZIP64_CENTRAL_DIRECTORY_LOCATOR), this.readBlockZip64EndOfCentralLocator(), !this.isSignature(this.relativeOffsetEndOfZip64CentralDir, o.ZIP64_CENTRAL_DIRECTORY_END) && (this.relativeOffsetEndOfZip64CentralDir = this.reader.lastIndexOfSignature(o.ZIP64_CENTRAL_DIRECTORY_END), this.relativeOffsetEndOfZip64CentralDir < 0)) throw new Error("Corrupted zip: can't find the ZIP64 end of central directory");
          this.reader.setIndex(this.relativeOffsetEndOfZip64CentralDir), this.checkSignature(o.ZIP64_CENTRAL_DIRECTORY_END), this.readBlockZip64EndOfCentral();
        }
        var c = this.centralDirOffset + this.centralDirSize;
        this.zip64 && (c += 20, c += 12 + this.zip64EndOfCentralSize);
        var g = u - c;
        if (0 < g) this.isSignature(u, o.CENTRAL_FILE_HEADER) || (this.reader.zero = g);
        else if (g < 0) throw new Error("Corrupted zip: missing " + Math.abs(g) + " bytes.");
      }, prepareReader: function(d) {
        this.reader = a(d);
      }, load: function(d) {
        this.prepareReader(d), this.readEndOfCentral(), this.readCentralDir(), this.readLocalFiles();
      } }, n.exports = f;
    }, { "./reader/readerFor": 22, "./signature": 23, "./support": 30, "./utils": 32, "./zipEntry": 34 }], 34: [function(e, n, i) {
      var a = e("./reader/readerFor"), s = e("./utils"), o = e("./compressedObject"), h = e("./crc32"), l = e("./utf8"), f = e("./compressions"), d = e("./support");
      function u(c, g) {
        this.options = c, this.loadOptions = g;
      }
      u.prototype = { isEncrypted: function() {
        return (1 & this.bitFlag) == 1;
      }, useUTF8: function() {
        return (2048 & this.bitFlag) == 2048;
      }, readLocalPart: function(c) {
        var g, p;
        if (c.skip(22), this.fileNameLength = c.readInt(2), p = c.readInt(2), this.fileName = c.readData(this.fileNameLength), c.skip(p), this.compressedSize === -1 || this.uncompressedSize === -1) throw new Error("Bug or corrupted zip : didn't get enough information from the central directory (compressedSize === -1 || uncompressedSize === -1)");
        if ((g = function(y) {
          for (var m in f) if (Object.prototype.hasOwnProperty.call(f, m) && f[m].magic === y) return f[m];
          return null;
        }(this.compressionMethod)) === null) throw new Error("Corrupted zip : compression " + s.pretty(this.compressionMethod) + " unknown (inner file : " + s.transformTo("string", this.fileName) + ")");
        this.decompressed = new o(this.compressedSize, this.uncompressedSize, this.crc32, g, c.readData(this.compressedSize));
      }, readCentralPart: function(c) {
        this.versionMadeBy = c.readInt(2), c.skip(2), this.bitFlag = c.readInt(2), this.compressionMethod = c.readString(2), this.date = c.readDate(), this.crc32 = c.readInt(4), this.compressedSize = c.readInt(4), this.uncompressedSize = c.readInt(4);
        var g = c.readInt(2);
        if (this.extraFieldsLength = c.readInt(2), this.fileCommentLength = c.readInt(2), this.diskNumberStart = c.readInt(2), this.internalFileAttributes = c.readInt(2), this.externalFileAttributes = c.readInt(4), this.localHeaderOffset = c.readInt(4), this.isEncrypted()) throw new Error("Encrypted zip are not supported");
        c.skip(g), this.readExtraFields(c), this.parseZIP64ExtraField(c), this.fileComment = c.readData(this.fileCommentLength);
      }, processAttributes: function() {
        this.unixPermissions = null, this.dosPermissions = null;
        var c = this.versionMadeBy >> 8;
        this.dir = !!(16 & this.externalFileAttributes), c == 0 && (this.dosPermissions = 63 & this.externalFileAttributes), c == 3 && (this.unixPermissions = this.externalFileAttributes >> 16 & 65535), this.dir || this.fileNameStr.slice(-1) !== "/" || (this.dir = !0);
      }, parseZIP64ExtraField: function() {
        if (this.extraFields[1]) {
          var c = a(this.extraFields[1].value);
          this.uncompressedSize === s.MAX_VALUE_32BITS && (this.uncompressedSize = c.readInt(8)), this.compressedSize === s.MAX_VALUE_32BITS && (this.compressedSize = c.readInt(8)), this.localHeaderOffset === s.MAX_VALUE_32BITS && (this.localHeaderOffset = c.readInt(8)), this.diskNumberStart === s.MAX_VALUE_32BITS && (this.diskNumberStart = c.readInt(4));
        }
      }, readExtraFields: function(c) {
        var g, p, y, m = c.index + this.extraFieldsLength;
        for (this.extraFields || (this.extraFields = {}); c.index + 4 < m; ) g = c.readInt(2), p = c.readInt(2), y = c.readData(p), this.extraFields[g] = { id: g, length: p, value: y };
        c.setIndex(m);
      }, handleUTF8: function() {
        var c = d.uint8array ? "uint8array" : "array";
        if (this.useUTF8()) this.fileNameStr = l.utf8decode(this.fileName), this.fileCommentStr = l.utf8decode(this.fileComment);
        else {
          var g = this.findExtraFieldUnicodePath();
          if (g !== null) this.fileNameStr = g;
          else {
            var p = s.transformTo(c, this.fileName);
            this.fileNameStr = this.loadOptions.decodeFileName(p);
          }
          var y = this.findExtraFieldUnicodeComment();
          if (y !== null) this.fileCommentStr = y;
          else {
            var m = s.transformTo(c, this.fileComment);
            this.fileCommentStr = this.loadOptions.decodeFileName(m);
          }
        }
      }, findExtraFieldUnicodePath: function() {
        var c = this.extraFields[28789];
        if (c) {
          var g = a(c.value);
          return g.readInt(1) !== 1 || h(this.fileName) !== g.readInt(4) ? null : l.utf8decode(g.readData(c.length - 5));
        }
        return null;
      }, findExtraFieldUnicodeComment: function() {
        var c = this.extraFields[25461];
        if (c) {
          var g = a(c.value);
          return g.readInt(1) !== 1 || h(this.fileComment) !== g.readInt(4) ? null : l.utf8decode(g.readData(c.length - 5));
        }
        return null;
      } }, n.exports = u;
    }, { "./compressedObject": 2, "./compressions": 3, "./crc32": 4, "./reader/readerFor": 22, "./support": 30, "./utf8": 31, "./utils": 32 }], 35: [function(e, n, i) {
      function a(g, p, y) {
        this.name = g, this.dir = y.dir, this.date = y.date, this.comment = y.comment, this.unixPermissions = y.unixPermissions, this.dosPermissions = y.dosPermissions, this._data = p, this._dataBinary = y.binary, this.options = { compression: y.compression, compressionOptions: y.compressionOptions };
      }
      var s = e("./stream/StreamHelper"), o = e("./stream/DataWorker"), h = e("./utf8"), l = e("./compressedObject"), f = e("./stream/GenericWorker");
      a.prototype = { internalStream: function(g) {
        var p = null, y = "string";
        try {
          if (!g) throw new Error("No output type specified.");
          var m = (y = g.toLowerCase()) === "string" || y === "text";
          y !== "binarystring" && y !== "text" || (y = "string"), p = this._decompressWorker();
          var A = !this._dataBinary;
          A && !m && (p = p.pipe(new h.Utf8EncodeWorker())), !A && m && (p = p.pipe(new h.Utf8DecodeWorker()));
        } catch (E) {
          (p = new f("error")).error(E);
        }
        return new s(p, y, "");
      }, async: function(g, p) {
        return this.internalStream(g).accumulate(p);
      }, nodeStream: function(g, p) {
        return this.internalStream(g || "nodebuffer").toNodejsStream(p);
      }, _compressWorker: function(g, p) {
        if (this._data instanceof l && this._data.compression.magic === g.magic) return this._data.getCompressedWorker();
        var y = this._decompressWorker();
        return this._dataBinary || (y = y.pipe(new h.Utf8EncodeWorker())), l.createWorkerFrom(y, g, p);
      }, _decompressWorker: function() {
        return this._data instanceof l ? this._data.getContentWorker() : this._data instanceof f ? this._data : new o(this._data);
      } };
      for (var d = ["asText", "asBinary", "asNodeBuffer", "asUint8Array", "asArrayBuffer"], u = function() {
        throw new Error("This method has been removed in JSZip 3.0, please check the upgrade guide.");
      }, c = 0; c < d.length; c++) a.prototype[d[c]] = u;
      n.exports = a;
    }, { "./compressedObject": 2, "./stream/DataWorker": 27, "./stream/GenericWorker": 28, "./stream/StreamHelper": 29, "./utf8": 31 }], 36: [function(e, n, i) {
      (function(a) {
        var s, o, h = a.MutationObserver || a.WebKitMutationObserver;
        if (h) {
          var l = 0, f = new h(g), d = a.document.createTextNode("");
          f.observe(d, { characterData: !0 }), s = function() {
            d.data = l = ++l % 2;
          };
        } else if (a.setImmediate || a.MessageChannel === void 0) s = "document" in a && "onreadystatechange" in a.document.createElement("script") ? function() {
          var p = a.document.createElement("script");
          p.onreadystatechange = function() {
            g(), p.onreadystatechange = null, p.parentNode.removeChild(p), p = null;
          }, a.document.documentElement.appendChild(p);
        } : function() {
          setTimeout(g, 0);
        };
        else {
          var u = new a.MessageChannel();
          u.port1.onmessage = g, s = function() {
            u.port2.postMessage(0);
          };
        }
        var c = [];
        function g() {
          var p, y;
          o = !0;
          for (var m = c.length; m; ) {
            for (y = c, c = [], p = -1; ++p < m; ) y[p]();
            m = c.length;
          }
          o = !1;
        }
        n.exports = function(p) {
          c.push(p) !== 1 || o || s();
        };
      }).call(this, typeof mr < "u" ? mr : typeof self < "u" ? self : typeof window < "u" ? window : {});
    }, {}], 37: [function(e, n, i) {
      var a = e("immediate");
      function s() {
      }
      var o = {}, h = ["REJECTED"], l = ["FULFILLED"], f = ["PENDING"];
      function d(m) {
        if (typeof m != "function") throw new TypeError("resolver must be a function");
        this.state = f, this.queue = [], this.outcome = void 0, m !== s && p(this, m);
      }
      function u(m, A, E) {
        this.promise = m, typeof A == "function" && (this.onFulfilled = A, this.callFulfilled = this.otherCallFulfilled), typeof E == "function" && (this.onRejected = E, this.callRejected = this.otherCallRejected);
      }
      function c(m, A, E) {
        a(function() {
          var x;
          try {
            x = A(E);
          } catch (S) {
            return o.reject(m, S);
          }
          x === m ? o.reject(m, new TypeError("Cannot resolve promise with itself")) : o.resolve(m, x);
        });
      }
      function g(m) {
        var A = m && m.then;
        if (m && (typeof m == "object" || typeof m == "function") && typeof A == "function") return function() {
          A.apply(m, arguments);
        };
      }
      function p(m, A) {
        var E = !1;
        function x(v) {
          E || (E = !0, o.reject(m, v));
        }
        function S(v) {
          E || (E = !0, o.resolve(m, v));
        }
        var R = y(function() {
          A(S, x);
        });
        R.status === "error" && x(R.value);
      }
      function y(m, A) {
        var E = {};
        try {
          E.value = m(A), E.status = "success";
        } catch (x) {
          E.status = "error", E.value = x;
        }
        return E;
      }
      (n.exports = d).prototype.finally = function(m) {
        if (typeof m != "function") return this;
        var A = this.constructor;
        return this.then(function(E) {
          return A.resolve(m()).then(function() {
            return E;
          });
        }, function(E) {
          return A.resolve(m()).then(function() {
            throw E;
          });
        });
      }, d.prototype.catch = function(m) {
        return this.then(null, m);
      }, d.prototype.then = function(m, A) {
        if (typeof m != "function" && this.state === l || typeof A != "function" && this.state === h) return this;
        var E = new this.constructor(s);
        return this.state !== f ? c(E, this.state === l ? m : A, this.outcome) : this.queue.push(new u(E, m, A)), E;
      }, u.prototype.callFulfilled = function(m) {
        o.resolve(this.promise, m);
      }, u.prototype.otherCallFulfilled = function(m) {
        c(this.promise, this.onFulfilled, m);
      }, u.prototype.callRejected = function(m) {
        o.reject(this.promise, m);
      }, u.prototype.otherCallRejected = function(m) {
        c(this.promise, this.onRejected, m);
      }, o.resolve = function(m, A) {
        var E = y(g, A);
        if (E.status === "error") return o.reject(m, E.value);
        var x = E.value;
        if (x) p(m, x);
        else {
          m.state = l, m.outcome = A;
          for (var S = -1, R = m.queue.length; ++S < R; ) m.queue[S].callFulfilled(A);
        }
        return m;
      }, o.reject = function(m, A) {
        m.state = h, m.outcome = A;
        for (var E = -1, x = m.queue.length; ++E < x; ) m.queue[E].callRejected(A);
        return m;
      }, d.resolve = function(m) {
        return m instanceof this ? m : o.resolve(new this(s), m);
      }, d.reject = function(m) {
        var A = new this(s);
        return o.reject(A, m);
      }, d.all = function(m) {
        var A = this;
        if (Object.prototype.toString.call(m) !== "[object Array]") return this.reject(new TypeError("must be an array"));
        var E = m.length, x = !1;
        if (!E) return this.resolve([]);
        for (var S = new Array(E), R = 0, v = -1, O = new this(s); ++v < E; ) C(m[v], v);
        return O;
        function C(B, Z) {
          A.resolve(B).then(function(T) {
            S[Z] = T, ++R !== E || x || (x = !0, o.resolve(O, S));
          }, function(T) {
            x || (x = !0, o.reject(O, T));
          });
        }
      }, d.race = function(m) {
        var A = this;
        if (Object.prototype.toString.call(m) !== "[object Array]") return this.reject(new TypeError("must be an array"));
        var E = m.length, x = !1;
        if (!E) return this.resolve([]);
        for (var S = -1, R = new this(s); ++S < E; ) v = m[S], A.resolve(v).then(function(O) {
          x || (x = !0, o.resolve(R, O));
        }, function(O) {
          x || (x = !0, o.reject(R, O));
        });
        var v;
        return R;
      };
    }, { immediate: 36 }], 38: [function(e, n, i) {
      var a = {};
      (0, e("./lib/utils/common").assign)(a, e("./lib/deflate"), e("./lib/inflate"), e("./lib/zlib/constants")), n.exports = a;
    }, { "./lib/deflate": 39, "./lib/inflate": 40, "./lib/utils/common": 41, "./lib/zlib/constants": 44 }], 39: [function(e, n, i) {
      var a = e("./zlib/deflate"), s = e("./utils/common"), o = e("./utils/strings"), h = e("./zlib/messages"), l = e("./zlib/zstream"), f = Object.prototype.toString, d = 0, u = -1, c = 0, g = 8;
      function p(m) {
        if (!(this instanceof p)) return new p(m);
        this.options = s.assign({ level: u, method: g, chunkSize: 16384, windowBits: 15, memLevel: 8, strategy: c, to: "" }, m || {});
        var A = this.options;
        A.raw && 0 < A.windowBits ? A.windowBits = -A.windowBits : A.gzip && 0 < A.windowBits && A.windowBits < 16 && (A.windowBits += 16), this.err = 0, this.msg = "", this.ended = !1, this.chunks = [], this.strm = new l(), this.strm.avail_out = 0;
        var E = a.deflateInit2(this.strm, A.level, A.method, A.windowBits, A.memLevel, A.strategy);
        if (E !== d) throw new Error(h[E]);
        if (A.header && a.deflateSetHeader(this.strm, A.header), A.dictionary) {
          var x;
          if (x = typeof A.dictionary == "string" ? o.string2buf(A.dictionary) : f.call(A.dictionary) === "[object ArrayBuffer]" ? new Uint8Array(A.dictionary) : A.dictionary, (E = a.deflateSetDictionary(this.strm, x)) !== d) throw new Error(h[E]);
          this._dict_set = !0;
        }
      }
      function y(m, A) {
        var E = new p(A);
        if (E.push(m, !0), E.err) throw E.msg || h[E.err];
        return E.result;
      }
      p.prototype.push = function(m, A) {
        var E, x, S = this.strm, R = this.options.chunkSize;
        if (this.ended) return !1;
        x = A === ~~A ? A : A === !0 ? 4 : 0, typeof m == "string" ? S.input = o.string2buf(m) : f.call(m) === "[object ArrayBuffer]" ? S.input = new Uint8Array(m) : S.input = m, S.next_in = 0, S.avail_in = S.input.length;
        do {
          if (S.avail_out === 0 && (S.output = new s.Buf8(R), S.next_out = 0, S.avail_out = R), (E = a.deflate(S, x)) !== 1 && E !== d) return this.onEnd(E), !(this.ended = !0);
          S.avail_out !== 0 && (S.avail_in !== 0 || x !== 4 && x !== 2) || (this.options.to === "string" ? this.onData(o.buf2binstring(s.shrinkBuf(S.output, S.next_out))) : this.onData(s.shrinkBuf(S.output, S.next_out)));
        } while ((0 < S.avail_in || S.avail_out === 0) && E !== 1);
        return x === 4 ? (E = a.deflateEnd(this.strm), this.onEnd(E), this.ended = !0, E === d) : x !== 2 || (this.onEnd(d), !(S.avail_out = 0));
      }, p.prototype.onData = function(m) {
        this.chunks.push(m);
      }, p.prototype.onEnd = function(m) {
        m === d && (this.options.to === "string" ? this.result = this.chunks.join("") : this.result = s.flattenChunks(this.chunks)), this.chunks = [], this.err = m, this.msg = this.strm.msg;
      }, i.Deflate = p, i.deflate = y, i.deflateRaw = function(m, A) {
        return (A = A || {}).raw = !0, y(m, A);
      }, i.gzip = function(m, A) {
        return (A = A || {}).gzip = !0, y(m, A);
      };
    }, { "./utils/common": 41, "./utils/strings": 42, "./zlib/deflate": 46, "./zlib/messages": 51, "./zlib/zstream": 53 }], 40: [function(e, n, i) {
      var a = e("./zlib/inflate"), s = e("./utils/common"), o = e("./utils/strings"), h = e("./zlib/constants"), l = e("./zlib/messages"), f = e("./zlib/zstream"), d = e("./zlib/gzheader"), u = Object.prototype.toString;
      function c(p) {
        if (!(this instanceof c)) return new c(p);
        this.options = s.assign({ chunkSize: 16384, windowBits: 0, to: "" }, p || {});
        var y = this.options;
        y.raw && 0 <= y.windowBits && y.windowBits < 16 && (y.windowBits = -y.windowBits, y.windowBits === 0 && (y.windowBits = -15)), !(0 <= y.windowBits && y.windowBits < 16) || p && p.windowBits || (y.windowBits += 32), 15 < y.windowBits && y.windowBits < 48 && !(15 & y.windowBits) && (y.windowBits |= 15), this.err = 0, this.msg = "", this.ended = !1, this.chunks = [], this.strm = new f(), this.strm.avail_out = 0;
        var m = a.inflateInit2(this.strm, y.windowBits);
        if (m !== h.Z_OK) throw new Error(l[m]);
        this.header = new d(), a.inflateGetHeader(this.strm, this.header);
      }
      function g(p, y) {
        var m = new c(y);
        if (m.push(p, !0), m.err) throw m.msg || l[m.err];
        return m.result;
      }
      c.prototype.push = function(p, y) {
        var m, A, E, x, S, R, v = this.strm, O = this.options.chunkSize, C = this.options.dictionary, B = !1;
        if (this.ended) return !1;
        A = y === ~~y ? y : y === !0 ? h.Z_FINISH : h.Z_NO_FLUSH, typeof p == "string" ? v.input = o.binstring2buf(p) : u.call(p) === "[object ArrayBuffer]" ? v.input = new Uint8Array(p) : v.input = p, v.next_in = 0, v.avail_in = v.input.length;
        do {
          if (v.avail_out === 0 && (v.output = new s.Buf8(O), v.next_out = 0, v.avail_out = O), (m = a.inflate(v, h.Z_NO_FLUSH)) === h.Z_NEED_DICT && C && (R = typeof C == "string" ? o.string2buf(C) : u.call(C) === "[object ArrayBuffer]" ? new Uint8Array(C) : C, m = a.inflateSetDictionary(this.strm, R)), m === h.Z_BUF_ERROR && B === !0 && (m = h.Z_OK, B = !1), m !== h.Z_STREAM_END && m !== h.Z_OK) return this.onEnd(m), !(this.ended = !0);
          v.next_out && (v.avail_out !== 0 && m !== h.Z_STREAM_END && (v.avail_in !== 0 || A !== h.Z_FINISH && A !== h.Z_SYNC_FLUSH) || (this.options.to === "string" ? (E = o.utf8border(v.output, v.next_out), x = v.next_out - E, S = o.buf2string(v.output, E), v.next_out = x, v.avail_out = O - x, x && s.arraySet(v.output, v.output, E, x, 0), this.onData(S)) : this.onData(s.shrinkBuf(v.output, v.next_out)))), v.avail_in === 0 && v.avail_out === 0 && (B = !0);
        } while ((0 < v.avail_in || v.avail_out === 0) && m !== h.Z_STREAM_END);
        return m === h.Z_STREAM_END && (A = h.Z_FINISH), A === h.Z_FINISH ? (m = a.inflateEnd(this.strm), this.onEnd(m), this.ended = !0, m === h.Z_OK) : A !== h.Z_SYNC_FLUSH || (this.onEnd(h.Z_OK), !(v.avail_out = 0));
      }, c.prototype.onData = function(p) {
        this.chunks.push(p);
      }, c.prototype.onEnd = function(p) {
        p === h.Z_OK && (this.options.to === "string" ? this.result = this.chunks.join("") : this.result = s.flattenChunks(this.chunks)), this.chunks = [], this.err = p, this.msg = this.strm.msg;
      }, i.Inflate = c, i.inflate = g, i.inflateRaw = function(p, y) {
        return (y = y || {}).raw = !0, g(p, y);
      }, i.ungzip = g;
    }, { "./utils/common": 41, "./utils/strings": 42, "./zlib/constants": 44, "./zlib/gzheader": 47, "./zlib/inflate": 49, "./zlib/messages": 51, "./zlib/zstream": 53 }], 41: [function(e, n, i) {
      var a = typeof Uint8Array < "u" && typeof Uint16Array < "u" && typeof Int32Array < "u";
      i.assign = function(h) {
        for (var l = Array.prototype.slice.call(arguments, 1); l.length; ) {
          var f = l.shift();
          if (f) {
            if (typeof f != "object") throw new TypeError(f + "must be non-object");
            for (var d in f) f.hasOwnProperty(d) && (h[d] = f[d]);
          }
        }
        return h;
      }, i.shrinkBuf = function(h, l) {
        return h.length === l ? h : h.subarray ? h.subarray(0, l) : (h.length = l, h);
      };
      var s = { arraySet: function(h, l, f, d, u) {
        if (l.subarray && h.subarray) h.set(l.subarray(f, f + d), u);
        else for (var c = 0; c < d; c++) h[u + c] = l[f + c];
      }, flattenChunks: function(h) {
        var l, f, d, u, c, g;
        for (l = d = 0, f = h.length; l < f; l++) d += h[l].length;
        for (g = new Uint8Array(d), l = u = 0, f = h.length; l < f; l++) c = h[l], g.set(c, u), u += c.length;
        return g;
      } }, o = { arraySet: function(h, l, f, d, u) {
        for (var c = 0; c < d; c++) h[u + c] = l[f + c];
      }, flattenChunks: function(h) {
        return [].concat.apply([], h);
      } };
      i.setTyped = function(h) {
        h ? (i.Buf8 = Uint8Array, i.Buf16 = Uint16Array, i.Buf32 = Int32Array, i.assign(i, s)) : (i.Buf8 = Array, i.Buf16 = Array, i.Buf32 = Array, i.assign(i, o));
      }, i.setTyped(a);
    }, {}], 42: [function(e, n, i) {
      var a = e("./common"), s = !0, o = !0;
      try {
        String.fromCharCode.apply(null, [0]);
      } catch {
        s = !1;
      }
      try {
        String.fromCharCode.apply(null, new Uint8Array(1));
      } catch {
        o = !1;
      }
      for (var h = new a.Buf8(256), l = 0; l < 256; l++) h[l] = 252 <= l ? 6 : 248 <= l ? 5 : 240 <= l ? 4 : 224 <= l ? 3 : 192 <= l ? 2 : 1;
      function f(d, u) {
        if (u < 65537 && (d.subarray && o || !d.subarray && s)) return String.fromCharCode.apply(null, a.shrinkBuf(d, u));
        for (var c = "", g = 0; g < u; g++) c += String.fromCharCode(d[g]);
        return c;
      }
      h[254] = h[254] = 1, i.string2buf = function(d) {
        var u, c, g, p, y, m = d.length, A = 0;
        for (p = 0; p < m; p++) (64512 & (c = d.charCodeAt(p))) == 55296 && p + 1 < m && (64512 & (g = d.charCodeAt(p + 1))) == 56320 && (c = 65536 + (c - 55296 << 10) + (g - 56320), p++), A += c < 128 ? 1 : c < 2048 ? 2 : c < 65536 ? 3 : 4;
        for (u = new a.Buf8(A), p = y = 0; y < A; p++) (64512 & (c = d.charCodeAt(p))) == 55296 && p + 1 < m && (64512 & (g = d.charCodeAt(p + 1))) == 56320 && (c = 65536 + (c - 55296 << 10) + (g - 56320), p++), c < 128 ? u[y++] = c : (c < 2048 ? u[y++] = 192 | c >>> 6 : (c < 65536 ? u[y++] = 224 | c >>> 12 : (u[y++] = 240 | c >>> 18, u[y++] = 128 | c >>> 12 & 63), u[y++] = 128 | c >>> 6 & 63), u[y++] = 128 | 63 & c);
        return u;
      }, i.buf2binstring = function(d) {
        return f(d, d.length);
      }, i.binstring2buf = function(d) {
        for (var u = new a.Buf8(d.length), c = 0, g = u.length; c < g; c++) u[c] = d.charCodeAt(c);
        return u;
      }, i.buf2string = function(d, u) {
        var c, g, p, y, m = u || d.length, A = new Array(2 * m);
        for (c = g = 0; c < m; ) if ((p = d[c++]) < 128) A[g++] = p;
        else if (4 < (y = h[p])) A[g++] = 65533, c += y - 1;
        else {
          for (p &= y === 2 ? 31 : y === 3 ? 15 : 7; 1 < y && c < m; ) p = p << 6 | 63 & d[c++], y--;
          1 < y ? A[g++] = 65533 : p < 65536 ? A[g++] = p : (p -= 65536, A[g++] = 55296 | p >> 10 & 1023, A[g++] = 56320 | 1023 & p);
        }
        return f(A, g);
      }, i.utf8border = function(d, u) {
        var c;
        for ((u = u || d.length) > d.length && (u = d.length), c = u - 1; 0 <= c && (192 & d[c]) == 128; ) c--;
        return c < 0 || c === 0 ? u : c + h[d[c]] > u ? c : u;
      };
    }, { "./common": 41 }], 43: [function(e, n, i) {
      n.exports = function(a, s, o, h) {
        for (var l = 65535 & a | 0, f = a >>> 16 & 65535 | 0, d = 0; o !== 0; ) {
          for (o -= d = 2e3 < o ? 2e3 : o; f = f + (l = l + s[h++] | 0) | 0, --d; ) ;
          l %= 65521, f %= 65521;
        }
        return l | f << 16 | 0;
      };
    }, {}], 44: [function(e, n, i) {
      n.exports = { Z_NO_FLUSH: 0, Z_PARTIAL_FLUSH: 1, Z_SYNC_FLUSH: 2, Z_FULL_FLUSH: 3, Z_FINISH: 4, Z_BLOCK: 5, Z_TREES: 6, Z_OK: 0, Z_STREAM_END: 1, Z_NEED_DICT: 2, Z_ERRNO: -1, Z_STREAM_ERROR: -2, Z_DATA_ERROR: -3, Z_BUF_ERROR: -5, Z_NO_COMPRESSION: 0, Z_BEST_SPEED: 1, Z_BEST_COMPRESSION: 9, Z_DEFAULT_COMPRESSION: -1, Z_FILTERED: 1, Z_HUFFMAN_ONLY: 2, Z_RLE: 3, Z_FIXED: 4, Z_DEFAULT_STRATEGY: 0, Z_BINARY: 0, Z_TEXT: 1, Z_UNKNOWN: 2, Z_DEFLATED: 8 };
    }, {}], 45: [function(e, n, i) {
      var a = function() {
        for (var s, o = [], h = 0; h < 256; h++) {
          s = h;
          for (var l = 0; l < 8; l++) s = 1 & s ? 3988292384 ^ s >>> 1 : s >>> 1;
          o[h] = s;
        }
        return o;
      }();
      n.exports = function(s, o, h, l) {
        var f = a, d = l + h;
        s ^= -1;
        for (var u = l; u < d; u++) s = s >>> 8 ^ f[255 & (s ^ o[u])];
        return -1 ^ s;
      };
    }, {}], 46: [function(e, n, i) {
      var a, s = e("../utils/common"), o = e("./trees"), h = e("./adler32"), l = e("./crc32"), f = e("./messages"), d = 0, u = 4, c = 0, g = -2, p = -1, y = 4, m = 2, A = 8, E = 9, x = 286, S = 30, R = 19, v = 2 * x + 1, O = 15, C = 3, B = 258, Z = B + C + 1, T = 42, j = 113, _ = 1, K = 2, se = 3, P = 4;
      function G(w, ae) {
        return w.msg = f[ae], ae;
      }
      function U(w) {
        return (w << 1) - (4 < w ? 9 : 0);
      }
      function Q(w) {
        for (var ae = w.length; 0 <= --ae; ) w[ae] = 0;
      }
      function M(w) {
        var ae = w.state, ee = ae.pending;
        ee > w.avail_out && (ee = w.avail_out), ee !== 0 && (s.arraySet(w.output, ae.pending_buf, ae.pending_out, ee, w.next_out), w.next_out += ee, ae.pending_out += ee, w.total_out += ee, w.avail_out -= ee, ae.pending -= ee, ae.pending === 0 && (ae.pending_out = 0));
      }
      function X(w, ae) {
        o._tr_flush_block(w, 0 <= w.block_start ? w.block_start : -1, w.strstart - w.block_start, ae), w.block_start = w.strstart, M(w.strm);
      }
      function de(w, ae) {
        w.pending_buf[w.pending++] = ae;
      }
      function le(w, ae) {
        w.pending_buf[w.pending++] = ae >>> 8 & 255, w.pending_buf[w.pending++] = 255 & ae;
      }
      function ne(w, ae) {
        var ee, D, F = w.max_chain_length, z = w.strstart, W = w.prev_length, ie = w.nice_match, $ = w.strstart > w.w_size - Z ? w.strstart - (w.w_size - Z) : 0, re = w.window, he = w.w_mask, oe = w.prev, ge = w.strstart + B, Ie = re[z + W - 1], be = re[z + W];
        w.prev_length >= w.good_match && (F >>= 2), ie > w.lookahead && (ie = w.lookahead);
        do
          if (re[(ee = ae) + W] === be && re[ee + W - 1] === Ie && re[ee] === re[z] && re[++ee] === re[z + 1]) {
            z += 2, ee++;
            do
              ;
            while (re[++z] === re[++ee] && re[++z] === re[++ee] && re[++z] === re[++ee] && re[++z] === re[++ee] && re[++z] === re[++ee] && re[++z] === re[++ee] && re[++z] === re[++ee] && re[++z] === re[++ee] && z < ge);
            if (D = B - (ge - z), z = ge - B, W < D) {
              if (w.match_start = ae, ie <= (W = D)) break;
              Ie = re[z + W - 1], be = re[z + W];
            }
          }
        while ((ae = oe[ae & he]) > $ && --F != 0);
        return W <= w.lookahead ? W : w.lookahead;
      }
      function ye(w) {
        var ae, ee, D, F, z, W, ie, $, re, he, oe = w.w_size;
        do {
          if (F = w.window_size - w.lookahead - w.strstart, w.strstart >= oe + (oe - Z)) {
            for (s.arraySet(w.window, w.window, oe, oe, 0), w.match_start -= oe, w.strstart -= oe, w.block_start -= oe, ae = ee = w.hash_size; D = w.head[--ae], w.head[ae] = oe <= D ? D - oe : 0, --ee; ) ;
            for (ae = ee = oe; D = w.prev[--ae], w.prev[ae] = oe <= D ? D - oe : 0, --ee; ) ;
            F += oe;
          }
          if (w.strm.avail_in === 0) break;
          if (W = w.strm, ie = w.window, $ = w.strstart + w.lookahead, re = F, he = void 0, he = W.avail_in, re < he && (he = re), ee = he === 0 ? 0 : (W.avail_in -= he, s.arraySet(ie, W.input, W.next_in, he, $), W.state.wrap === 1 ? W.adler = h(W.adler, ie, he, $) : W.state.wrap === 2 && (W.adler = l(W.adler, ie, he, $)), W.next_in += he, W.total_in += he, he), w.lookahead += ee, w.lookahead + w.insert >= C) for (z = w.strstart - w.insert, w.ins_h = w.window[z], w.ins_h = (w.ins_h << w.hash_shift ^ w.window[z + 1]) & w.hash_mask; w.insert && (w.ins_h = (w.ins_h << w.hash_shift ^ w.window[z + C - 1]) & w.hash_mask, w.prev[z & w.w_mask] = w.head[w.ins_h], w.head[w.ins_h] = z, z++, w.insert--, !(w.lookahead + w.insert < C)); ) ;
        } while (w.lookahead < Z && w.strm.avail_in !== 0);
      }
      function Te(w, ae) {
        for (var ee, D; ; ) {
          if (w.lookahead < Z) {
            if (ye(w), w.lookahead < Z && ae === d) return _;
            if (w.lookahead === 0) break;
          }
          if (ee = 0, w.lookahead >= C && (w.ins_h = (w.ins_h << w.hash_shift ^ w.window[w.strstart + C - 1]) & w.hash_mask, ee = w.prev[w.strstart & w.w_mask] = w.head[w.ins_h], w.head[w.ins_h] = w.strstart), ee !== 0 && w.strstart - ee <= w.w_size - Z && (w.match_length = ne(w, ee)), w.match_length >= C) if (D = o._tr_tally(w, w.strstart - w.match_start, w.match_length - C), w.lookahead -= w.match_length, w.match_length <= w.max_lazy_match && w.lookahead >= C) {
            for (w.match_length--; w.strstart++, w.ins_h = (w.ins_h << w.hash_shift ^ w.window[w.strstart + C - 1]) & w.hash_mask, ee = w.prev[w.strstart & w.w_mask] = w.head[w.ins_h], w.head[w.ins_h] = w.strstart, --w.match_length != 0; ) ;
            w.strstart++;
          } else w.strstart += w.match_length, w.match_length = 0, w.ins_h = w.window[w.strstart], w.ins_h = (w.ins_h << w.hash_shift ^ w.window[w.strstart + 1]) & w.hash_mask;
          else D = o._tr_tally(w, 0, w.window[w.strstart]), w.lookahead--, w.strstart++;
          if (D && (X(w, !1), w.strm.avail_out === 0)) return _;
        }
        return w.insert = w.strstart < C - 1 ? w.strstart : C - 1, ae === u ? (X(w, !0), w.strm.avail_out === 0 ? se : P) : w.last_lit && (X(w, !1), w.strm.avail_out === 0) ? _ : K;
      }
      function _e(w, ae) {
        for (var ee, D, F; ; ) {
          if (w.lookahead < Z) {
            if (ye(w), w.lookahead < Z && ae === d) return _;
            if (w.lookahead === 0) break;
          }
          if (ee = 0, w.lookahead >= C && (w.ins_h = (w.ins_h << w.hash_shift ^ w.window[w.strstart + C - 1]) & w.hash_mask, ee = w.prev[w.strstart & w.w_mask] = w.head[w.ins_h], w.head[w.ins_h] = w.strstart), w.prev_length = w.match_length, w.prev_match = w.match_start, w.match_length = C - 1, ee !== 0 && w.prev_length < w.max_lazy_match && w.strstart - ee <= w.w_size - Z && (w.match_length = ne(w, ee), w.match_length <= 5 && (w.strategy === 1 || w.match_length === C && 4096 < w.strstart - w.match_start) && (w.match_length = C - 1)), w.prev_length >= C && w.match_length <= w.prev_length) {
            for (F = w.strstart + w.lookahead - C, D = o._tr_tally(w, w.strstart - 1 - w.prev_match, w.prev_length - C), w.lookahead -= w.prev_length - 1, w.prev_length -= 2; ++w.strstart <= F && (w.ins_h = (w.ins_h << w.hash_shift ^ w.window[w.strstart + C - 1]) & w.hash_mask, ee = w.prev[w.strstart & w.w_mask] = w.head[w.ins_h], w.head[w.ins_h] = w.strstart), --w.prev_length != 0; ) ;
            if (w.match_available = 0, w.match_length = C - 1, w.strstart++, D && (X(w, !1), w.strm.avail_out === 0)) return _;
          } else if (w.match_available) {
            if ((D = o._tr_tally(w, 0, w.window[w.strstart - 1])) && X(w, !1), w.strstart++, w.lookahead--, w.strm.avail_out === 0) return _;
          } else w.match_available = 1, w.strstart++, w.lookahead--;
        }
        return w.match_available && (D = o._tr_tally(w, 0, w.window[w.strstart - 1]), w.match_available = 0), w.insert = w.strstart < C - 1 ? w.strstart : C - 1, ae === u ? (X(w, !0), w.strm.avail_out === 0 ? se : P) : w.last_lit && (X(w, !1), w.strm.avail_out === 0) ? _ : K;
      }
      function Se(w, ae, ee, D, F) {
        this.good_length = w, this.max_lazy = ae, this.nice_length = ee, this.max_chain = D, this.func = F;
      }
      function ve() {
        this.strm = null, this.status = 0, this.pending_buf = null, this.pending_buf_size = 0, this.pending_out = 0, this.pending = 0, this.wrap = 0, this.gzhead = null, this.gzindex = 0, this.method = A, this.last_flush = -1, this.w_size = 0, this.w_bits = 0, this.w_mask = 0, this.window = null, this.window_size = 0, this.prev = null, this.head = null, this.ins_h = 0, this.hash_size = 0, this.hash_bits = 0, this.hash_mask = 0, this.hash_shift = 0, this.block_start = 0, this.match_length = 0, this.prev_match = 0, this.match_available = 0, this.strstart = 0, this.match_start = 0, this.lookahead = 0, this.prev_length = 0, this.max_chain_length = 0, this.max_lazy_match = 0, this.level = 0, this.strategy = 0, this.good_match = 0, this.nice_match = 0, this.dyn_ltree = new s.Buf16(2 * v), this.dyn_dtree = new s.Buf16(2 * (2 * S + 1)), this.bl_tree = new s.Buf16(2 * (2 * R + 1)), Q(this.dyn_ltree), Q(this.dyn_dtree), Q(this.bl_tree), this.l_desc = null, this.d_desc = null, this.bl_desc = null, this.bl_count = new s.Buf16(O + 1), this.heap = new s.Buf16(2 * x + 1), Q(this.heap), this.heap_len = 0, this.heap_max = 0, this.depth = new s.Buf16(2 * x + 1), Q(this.depth), this.l_buf = 0, this.lit_bufsize = 0, this.last_lit = 0, this.d_buf = 0, this.opt_len = 0, this.static_len = 0, this.matches = 0, this.insert = 0, this.bi_buf = 0, this.bi_valid = 0;
      }
      function Re(w) {
        var ae;
        return w && w.state ? (w.total_in = w.total_out = 0, w.data_type = m, (ae = w.state).pending = 0, ae.pending_out = 0, ae.wrap < 0 && (ae.wrap = -ae.wrap), ae.status = ae.wrap ? T : j, w.adler = ae.wrap === 2 ? 0 : 1, ae.last_flush = d, o._tr_init(ae), c) : G(w, g);
      }
      function Ue(w) {
        var ae = Re(w);
        return ae === c && function(ee) {
          ee.window_size = 2 * ee.w_size, Q(ee.head), ee.max_lazy_match = a[ee.level].max_lazy, ee.good_match = a[ee.level].good_length, ee.nice_match = a[ee.level].nice_length, ee.max_chain_length = a[ee.level].max_chain, ee.strstart = 0, ee.block_start = 0, ee.lookahead = 0, ee.insert = 0, ee.match_length = ee.prev_length = C - 1, ee.match_available = 0, ee.ins_h = 0;
        }(w.state), ae;
      }
      function He(w, ae, ee, D, F, z) {
        if (!w) return g;
        var W = 1;
        if (ae === p && (ae = 6), D < 0 ? (W = 0, D = -D) : 15 < D && (W = 2, D -= 16), F < 1 || E < F || ee !== A || D < 8 || 15 < D || ae < 0 || 9 < ae || z < 0 || y < z) return G(w, g);
        D === 8 && (D = 9);
        var ie = new ve();
        return (w.state = ie).strm = w, ie.wrap = W, ie.gzhead = null, ie.w_bits = D, ie.w_size = 1 << ie.w_bits, ie.w_mask = ie.w_size - 1, ie.hash_bits = F + 7, ie.hash_size = 1 << ie.hash_bits, ie.hash_mask = ie.hash_size - 1, ie.hash_shift = ~~((ie.hash_bits + C - 1) / C), ie.window = new s.Buf8(2 * ie.w_size), ie.head = new s.Buf16(ie.hash_size), ie.prev = new s.Buf16(ie.w_size), ie.lit_bufsize = 1 << F + 6, ie.pending_buf_size = 4 * ie.lit_bufsize, ie.pending_buf = new s.Buf8(ie.pending_buf_size), ie.d_buf = 1 * ie.lit_bufsize, ie.l_buf = 3 * ie.lit_bufsize, ie.level = ae, ie.strategy = z, ie.method = ee, Ue(w);
      }
      a = [new Se(0, 0, 0, 0, function(w, ae) {
        var ee = 65535;
        for (ee > w.pending_buf_size - 5 && (ee = w.pending_buf_size - 5); ; ) {
          if (w.lookahead <= 1) {
            if (ye(w), w.lookahead === 0 && ae === d) return _;
            if (w.lookahead === 0) break;
          }
          w.strstart += w.lookahead, w.lookahead = 0;
          var D = w.block_start + ee;
          if ((w.strstart === 0 || w.strstart >= D) && (w.lookahead = w.strstart - D, w.strstart = D, X(w, !1), w.strm.avail_out === 0) || w.strstart - w.block_start >= w.w_size - Z && (X(w, !1), w.strm.avail_out === 0)) return _;
        }
        return w.insert = 0, ae === u ? (X(w, !0), w.strm.avail_out === 0 ? se : P) : (w.strstart > w.block_start && (X(w, !1), w.strm.avail_out), _);
      }), new Se(4, 4, 8, 4, Te), new Se(4, 5, 16, 8, Te), new Se(4, 6, 32, 32, Te), new Se(4, 4, 16, 16, _e), new Se(8, 16, 32, 32, _e), new Se(8, 16, 128, 128, _e), new Se(8, 32, 128, 256, _e), new Se(32, 128, 258, 1024, _e), new Se(32, 258, 258, 4096, _e)], i.deflateInit = function(w, ae) {
        return He(w, ae, A, 15, 8, 0);
      }, i.deflateInit2 = He, i.deflateReset = Ue, i.deflateResetKeep = Re, i.deflateSetHeader = function(w, ae) {
        return w && w.state ? w.state.wrap !== 2 ? g : (w.state.gzhead = ae, c) : g;
      }, i.deflate = function(w, ae) {
        var ee, D, F, z;
        if (!w || !w.state || 5 < ae || ae < 0) return w ? G(w, g) : g;
        if (D = w.state, !w.output || !w.input && w.avail_in !== 0 || D.status === 666 && ae !== u) return G(w, w.avail_out === 0 ? -5 : g);
        if (D.strm = w, ee = D.last_flush, D.last_flush = ae, D.status === T) if (D.wrap === 2) w.adler = 0, de(D, 31), de(D, 139), de(D, 8), D.gzhead ? (de(D, (D.gzhead.text ? 1 : 0) + (D.gzhead.hcrc ? 2 : 0) + (D.gzhead.extra ? 4 : 0) + (D.gzhead.name ? 8 : 0) + (D.gzhead.comment ? 16 : 0)), de(D, 255 & D.gzhead.time), de(D, D.gzhead.time >> 8 & 255), de(D, D.gzhead.time >> 16 & 255), de(D, D.gzhead.time >> 24 & 255), de(D, D.level === 9 ? 2 : 2 <= D.strategy || D.level < 2 ? 4 : 0), de(D, 255 & D.gzhead.os), D.gzhead.extra && D.gzhead.extra.length && (de(D, 255 & D.gzhead.extra.length), de(D, D.gzhead.extra.length >> 8 & 255)), D.gzhead.hcrc && (w.adler = l(w.adler, D.pending_buf, D.pending, 0)), D.gzindex = 0, D.status = 69) : (de(D, 0), de(D, 0), de(D, 0), de(D, 0), de(D, 0), de(D, D.level === 9 ? 2 : 2 <= D.strategy || D.level < 2 ? 4 : 0), de(D, 3), D.status = j);
        else {
          var W = A + (D.w_bits - 8 << 4) << 8;
          W |= (2 <= D.strategy || D.level < 2 ? 0 : D.level < 6 ? 1 : D.level === 6 ? 2 : 3) << 6, D.strstart !== 0 && (W |= 32), W += 31 - W % 31, D.status = j, le(D, W), D.strstart !== 0 && (le(D, w.adler >>> 16), le(D, 65535 & w.adler)), w.adler = 1;
        }
        if (D.status === 69) if (D.gzhead.extra) {
          for (F = D.pending; D.gzindex < (65535 & D.gzhead.extra.length) && (D.pending !== D.pending_buf_size || (D.gzhead.hcrc && D.pending > F && (w.adler = l(w.adler, D.pending_buf, D.pending - F, F)), M(w), F = D.pending, D.pending !== D.pending_buf_size)); ) de(D, 255 & D.gzhead.extra[D.gzindex]), D.gzindex++;
          D.gzhead.hcrc && D.pending > F && (w.adler = l(w.adler, D.pending_buf, D.pending - F, F)), D.gzindex === D.gzhead.extra.length && (D.gzindex = 0, D.status = 73);
        } else D.status = 73;
        if (D.status === 73) if (D.gzhead.name) {
          F = D.pending;
          do {
            if (D.pending === D.pending_buf_size && (D.gzhead.hcrc && D.pending > F && (w.adler = l(w.adler, D.pending_buf, D.pending - F, F)), M(w), F = D.pending, D.pending === D.pending_buf_size)) {
              z = 1;
              break;
            }
            z = D.gzindex < D.gzhead.name.length ? 255 & D.gzhead.name.charCodeAt(D.gzindex++) : 0, de(D, z);
          } while (z !== 0);
          D.gzhead.hcrc && D.pending > F && (w.adler = l(w.adler, D.pending_buf, D.pending - F, F)), z === 0 && (D.gzindex = 0, D.status = 91);
        } else D.status = 91;
        if (D.status === 91) if (D.gzhead.comment) {
          F = D.pending;
          do {
            if (D.pending === D.pending_buf_size && (D.gzhead.hcrc && D.pending > F && (w.adler = l(w.adler, D.pending_buf, D.pending - F, F)), M(w), F = D.pending, D.pending === D.pending_buf_size)) {
              z = 1;
              break;
            }
            z = D.gzindex < D.gzhead.comment.length ? 255 & D.gzhead.comment.charCodeAt(D.gzindex++) : 0, de(D, z);
          } while (z !== 0);
          D.gzhead.hcrc && D.pending > F && (w.adler = l(w.adler, D.pending_buf, D.pending - F, F)), z === 0 && (D.status = 103);
        } else D.status = 103;
        if (D.status === 103 && (D.gzhead.hcrc ? (D.pending + 2 > D.pending_buf_size && M(w), D.pending + 2 <= D.pending_buf_size && (de(D, 255 & w.adler), de(D, w.adler >> 8 & 255), w.adler = 0, D.status = j)) : D.status = j), D.pending !== 0) {
          if (M(w), w.avail_out === 0) return D.last_flush = -1, c;
        } else if (w.avail_in === 0 && U(ae) <= U(ee) && ae !== u) return G(w, -5);
        if (D.status === 666 && w.avail_in !== 0) return G(w, -5);
        if (w.avail_in !== 0 || D.lookahead !== 0 || ae !== d && D.status !== 666) {
          var ie = D.strategy === 2 ? function($, re) {
            for (var he; ; ) {
              if ($.lookahead === 0 && (ye($), $.lookahead === 0)) {
                if (re === d) return _;
                break;
              }
              if ($.match_length = 0, he = o._tr_tally($, 0, $.window[$.strstart]), $.lookahead--, $.strstart++, he && (X($, !1), $.strm.avail_out === 0)) return _;
            }
            return $.insert = 0, re === u ? (X($, !0), $.strm.avail_out === 0 ? se : P) : $.last_lit && (X($, !1), $.strm.avail_out === 0) ? _ : K;
          }(D, ae) : D.strategy === 3 ? function($, re) {
            for (var he, oe, ge, Ie, be = $.window; ; ) {
              if ($.lookahead <= B) {
                if (ye($), $.lookahead <= B && re === d) return _;
                if ($.lookahead === 0) break;
              }
              if ($.match_length = 0, $.lookahead >= C && 0 < $.strstart && (oe = be[ge = $.strstart - 1]) === be[++ge] && oe === be[++ge] && oe === be[++ge]) {
                Ie = $.strstart + B;
                do
                  ;
                while (oe === be[++ge] && oe === be[++ge] && oe === be[++ge] && oe === be[++ge] && oe === be[++ge] && oe === be[++ge] && oe === be[++ge] && oe === be[++ge] && ge < Ie);
                $.match_length = B - (Ie - ge), $.match_length > $.lookahead && ($.match_length = $.lookahead);
              }
              if ($.match_length >= C ? (he = o._tr_tally($, 1, $.match_length - C), $.lookahead -= $.match_length, $.strstart += $.match_length, $.match_length = 0) : (he = o._tr_tally($, 0, $.window[$.strstart]), $.lookahead--, $.strstart++), he && (X($, !1), $.strm.avail_out === 0)) return _;
            }
            return $.insert = 0, re === u ? (X($, !0), $.strm.avail_out === 0 ? se : P) : $.last_lit && (X($, !1), $.strm.avail_out === 0) ? _ : K;
          }(D, ae) : a[D.level].func(D, ae);
          if (ie !== se && ie !== P || (D.status = 666), ie === _ || ie === se) return w.avail_out === 0 && (D.last_flush = -1), c;
          if (ie === K && (ae === 1 ? o._tr_align(D) : ae !== 5 && (o._tr_stored_block(D, 0, 0, !1), ae === 3 && (Q(D.head), D.lookahead === 0 && (D.strstart = 0, D.block_start = 0, D.insert = 0))), M(w), w.avail_out === 0)) return D.last_flush = -1, c;
        }
        return ae !== u ? c : D.wrap <= 0 ? 1 : (D.wrap === 2 ? (de(D, 255 & w.adler), de(D, w.adler >> 8 & 255), de(D, w.adler >> 16 & 255), de(D, w.adler >> 24 & 255), de(D, 255 & w.total_in), de(D, w.total_in >> 8 & 255), de(D, w.total_in >> 16 & 255), de(D, w.total_in >> 24 & 255)) : (le(D, w.adler >>> 16), le(D, 65535 & w.adler)), M(w), 0 < D.wrap && (D.wrap = -D.wrap), D.pending !== 0 ? c : 1);
      }, i.deflateEnd = function(w) {
        var ae;
        return w && w.state ? (ae = w.state.status) !== T && ae !== 69 && ae !== 73 && ae !== 91 && ae !== 103 && ae !== j && ae !== 666 ? G(w, g) : (w.state = null, ae === j ? G(w, -3) : c) : g;
      }, i.deflateSetDictionary = function(w, ae) {
        var ee, D, F, z, W, ie, $, re, he = ae.length;
        if (!w || !w.state || (z = (ee = w.state).wrap) === 2 || z === 1 && ee.status !== T || ee.lookahead) return g;
        for (z === 1 && (w.adler = h(w.adler, ae, he, 0)), ee.wrap = 0, he >= ee.w_size && (z === 0 && (Q(ee.head), ee.strstart = 0, ee.block_start = 0, ee.insert = 0), re = new s.Buf8(ee.w_size), s.arraySet(re, ae, he - ee.w_size, ee.w_size, 0), ae = re, he = ee.w_size), W = w.avail_in, ie = w.next_in, $ = w.input, w.avail_in = he, w.next_in = 0, w.input = ae, ye(ee); ee.lookahead >= C; ) {
          for (D = ee.strstart, F = ee.lookahead - (C - 1); ee.ins_h = (ee.ins_h << ee.hash_shift ^ ee.window[D + C - 1]) & ee.hash_mask, ee.prev[D & ee.w_mask] = ee.head[ee.ins_h], ee.head[ee.ins_h] = D, D++, --F; ) ;
          ee.strstart = D, ee.lookahead = C - 1, ye(ee);
        }
        return ee.strstart += ee.lookahead, ee.block_start = ee.strstart, ee.insert = ee.lookahead, ee.lookahead = 0, ee.match_length = ee.prev_length = C - 1, ee.match_available = 0, w.next_in = ie, w.input = $, w.avail_in = W, ee.wrap = z, c;
      }, i.deflateInfo = "pako deflate (from Nodeca project)";
    }, { "../utils/common": 41, "./adler32": 43, "./crc32": 45, "./messages": 51, "./trees": 52 }], 47: [function(e, n, i) {
      n.exports = function() {
        this.text = 0, this.time = 0, this.xflags = 0, this.os = 0, this.extra = null, this.extra_len = 0, this.name = "", this.comment = "", this.hcrc = 0, this.done = !1;
      };
    }, {}], 48: [function(e, n, i) {
      n.exports = function(a, s) {
        var o, h, l, f, d, u, c, g, p, y, m, A, E, x, S, R, v, O, C, B, Z, T, j, _, K;
        o = a.state, h = a.next_in, _ = a.input, l = h + (a.avail_in - 5), f = a.next_out, K = a.output, d = f - (s - a.avail_out), u = f + (a.avail_out - 257), c = o.dmax, g = o.wsize, p = o.whave, y = o.wnext, m = o.window, A = o.hold, E = o.bits, x = o.lencode, S = o.distcode, R = (1 << o.lenbits) - 1, v = (1 << o.distbits) - 1;
        e: do {
          E < 15 && (A += _[h++] << E, E += 8, A += _[h++] << E, E += 8), O = x[A & R];
          t: for (; ; ) {
            if (A >>>= C = O >>> 24, E -= C, (C = O >>> 16 & 255) === 0) K[f++] = 65535 & O;
            else {
              if (!(16 & C)) {
                if (!(64 & C)) {
                  O = x[(65535 & O) + (A & (1 << C) - 1)];
                  continue t;
                }
                if (32 & C) {
                  o.mode = 12;
                  break e;
                }
                a.msg = "invalid literal/length code", o.mode = 30;
                break e;
              }
              B = 65535 & O, (C &= 15) && (E < C && (A += _[h++] << E, E += 8), B += A & (1 << C) - 1, A >>>= C, E -= C), E < 15 && (A += _[h++] << E, E += 8, A += _[h++] << E, E += 8), O = S[A & v];
              r: for (; ; ) {
                if (A >>>= C = O >>> 24, E -= C, !(16 & (C = O >>> 16 & 255))) {
                  if (!(64 & C)) {
                    O = S[(65535 & O) + (A & (1 << C) - 1)];
                    continue r;
                  }
                  a.msg = "invalid distance code", o.mode = 30;
                  break e;
                }
                if (Z = 65535 & O, E < (C &= 15) && (A += _[h++] << E, (E += 8) < C && (A += _[h++] << E, E += 8)), c < (Z += A & (1 << C) - 1)) {
                  a.msg = "invalid distance too far back", o.mode = 30;
                  break e;
                }
                if (A >>>= C, E -= C, (C = f - d) < Z) {
                  if (p < (C = Z - C) && o.sane) {
                    a.msg = "invalid distance too far back", o.mode = 30;
                    break e;
                  }
                  if (j = m, (T = 0) === y) {
                    if (T += g - C, C < B) {
                      for (B -= C; K[f++] = m[T++], --C; ) ;
                      T = f - Z, j = K;
                    }
                  } else if (y < C) {
                    if (T += g + y - C, (C -= y) < B) {
                      for (B -= C; K[f++] = m[T++], --C; ) ;
                      if (T = 0, y < B) {
                        for (B -= C = y; K[f++] = m[T++], --C; ) ;
                        T = f - Z, j = K;
                      }
                    }
                  } else if (T += y - C, C < B) {
                    for (B -= C; K[f++] = m[T++], --C; ) ;
                    T = f - Z, j = K;
                  }
                  for (; 2 < B; ) K[f++] = j[T++], K[f++] = j[T++], K[f++] = j[T++], B -= 3;
                  B && (K[f++] = j[T++], 1 < B && (K[f++] = j[T++]));
                } else {
                  for (T = f - Z; K[f++] = K[T++], K[f++] = K[T++], K[f++] = K[T++], 2 < (B -= 3); ) ;
                  B && (K[f++] = K[T++], 1 < B && (K[f++] = K[T++]));
                }
                break;
              }
            }
            break;
          }
        } while (h < l && f < u);
        h -= B = E >> 3, A &= (1 << (E -= B << 3)) - 1, a.next_in = h, a.next_out = f, a.avail_in = h < l ? l - h + 5 : 5 - (h - l), a.avail_out = f < u ? u - f + 257 : 257 - (f - u), o.hold = A, o.bits = E;
      };
    }, {}], 49: [function(e, n, i) {
      var a = e("../utils/common"), s = e("./adler32"), o = e("./crc32"), h = e("./inffast"), l = e("./inftrees"), f = 1, d = 2, u = 0, c = -2, g = 1, p = 852, y = 592;
      function m(T) {
        return (T >>> 24 & 255) + (T >>> 8 & 65280) + ((65280 & T) << 8) + ((255 & T) << 24);
      }
      function A() {
        this.mode = 0, this.last = !1, this.wrap = 0, this.havedict = !1, this.flags = 0, this.dmax = 0, this.check = 0, this.total = 0, this.head = null, this.wbits = 0, this.wsize = 0, this.whave = 0, this.wnext = 0, this.window = null, this.hold = 0, this.bits = 0, this.length = 0, this.offset = 0, this.extra = 0, this.lencode = null, this.distcode = null, this.lenbits = 0, this.distbits = 0, this.ncode = 0, this.nlen = 0, this.ndist = 0, this.have = 0, this.next = null, this.lens = new a.Buf16(320), this.work = new a.Buf16(288), this.lendyn = null, this.distdyn = null, this.sane = 0, this.back = 0, this.was = 0;
      }
      function E(T) {
        var j;
        return T && T.state ? (j = T.state, T.total_in = T.total_out = j.total = 0, T.msg = "", j.wrap && (T.adler = 1 & j.wrap), j.mode = g, j.last = 0, j.havedict = 0, j.dmax = 32768, j.head = null, j.hold = 0, j.bits = 0, j.lencode = j.lendyn = new a.Buf32(p), j.distcode = j.distdyn = new a.Buf32(y), j.sane = 1, j.back = -1, u) : c;
      }
      function x(T) {
        var j;
        return T && T.state ? ((j = T.state).wsize = 0, j.whave = 0, j.wnext = 0, E(T)) : c;
      }
      function S(T, j) {
        var _, K;
        return T && T.state ? (K = T.state, j < 0 ? (_ = 0, j = -j) : (_ = 1 + (j >> 4), j < 48 && (j &= 15)), j && (j < 8 || 15 < j) ? c : (K.window !== null && K.wbits !== j && (K.window = null), K.wrap = _, K.wbits = j, x(T))) : c;
      }
      function R(T, j) {
        var _, K;
        return T ? (K = new A(), (T.state = K).window = null, (_ = S(T, j)) !== u && (T.state = null), _) : c;
      }
      var v, O, C = !0;
      function B(T) {
        if (C) {
          var j;
          for (v = new a.Buf32(512), O = new a.Buf32(32), j = 0; j < 144; ) T.lens[j++] = 8;
          for (; j < 256; ) T.lens[j++] = 9;
          for (; j < 280; ) T.lens[j++] = 7;
          for (; j < 288; ) T.lens[j++] = 8;
          for (l(f, T.lens, 0, 288, v, 0, T.work, { bits: 9 }), j = 0; j < 32; ) T.lens[j++] = 5;
          l(d, T.lens, 0, 32, O, 0, T.work, { bits: 5 }), C = !1;
        }
        T.lencode = v, T.lenbits = 9, T.distcode = O, T.distbits = 5;
      }
      function Z(T, j, _, K) {
        var se, P = T.state;
        return P.window === null && (P.wsize = 1 << P.wbits, P.wnext = 0, P.whave = 0, P.window = new a.Buf8(P.wsize)), K >= P.wsize ? (a.arraySet(P.window, j, _ - P.wsize, P.wsize, 0), P.wnext = 0, P.whave = P.wsize) : (K < (se = P.wsize - P.wnext) && (se = K), a.arraySet(P.window, j, _ - K, se, P.wnext), (K -= se) ? (a.arraySet(P.window, j, _ - K, K, 0), P.wnext = K, P.whave = P.wsize) : (P.wnext += se, P.wnext === P.wsize && (P.wnext = 0), P.whave < P.wsize && (P.whave += se))), 0;
      }
      i.inflateReset = x, i.inflateReset2 = S, i.inflateResetKeep = E, i.inflateInit = function(T) {
        return R(T, 15);
      }, i.inflateInit2 = R, i.inflate = function(T, j) {
        var _, K, se, P, G, U, Q, M, X, de, le, ne, ye, Te, _e, Se, ve, Re, Ue, He, w, ae, ee, D, F = 0, z = new a.Buf8(4), W = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];
        if (!T || !T.state || !T.output || !T.input && T.avail_in !== 0) return c;
        (_ = T.state).mode === 12 && (_.mode = 13), G = T.next_out, se = T.output, Q = T.avail_out, P = T.next_in, K = T.input, U = T.avail_in, M = _.hold, X = _.bits, de = U, le = Q, ae = u;
        e: for (; ; ) switch (_.mode) {
          case g:
            if (_.wrap === 0) {
              _.mode = 13;
              break;
            }
            for (; X < 16; ) {
              if (U === 0) break e;
              U--, M += K[P++] << X, X += 8;
            }
            if (2 & _.wrap && M === 35615) {
              z[_.check = 0] = 255 & M, z[1] = M >>> 8 & 255, _.check = o(_.check, z, 2, 0), X = M = 0, _.mode = 2;
              break;
            }
            if (_.flags = 0, _.head && (_.head.done = !1), !(1 & _.wrap) || (((255 & M) << 8) + (M >> 8)) % 31) {
              T.msg = "incorrect header check", _.mode = 30;
              break;
            }
            if ((15 & M) != 8) {
              T.msg = "unknown compression method", _.mode = 30;
              break;
            }
            if (X -= 4, w = 8 + (15 & (M >>>= 4)), _.wbits === 0) _.wbits = w;
            else if (w > _.wbits) {
              T.msg = "invalid window size", _.mode = 30;
              break;
            }
            _.dmax = 1 << w, T.adler = _.check = 1, _.mode = 512 & M ? 10 : 12, X = M = 0;
            break;
          case 2:
            for (; X < 16; ) {
              if (U === 0) break e;
              U--, M += K[P++] << X, X += 8;
            }
            if (_.flags = M, (255 & _.flags) != 8) {
              T.msg = "unknown compression method", _.mode = 30;
              break;
            }
            if (57344 & _.flags) {
              T.msg = "unknown header flags set", _.mode = 30;
              break;
            }
            _.head && (_.head.text = M >> 8 & 1), 512 & _.flags && (z[0] = 255 & M, z[1] = M >>> 8 & 255, _.check = o(_.check, z, 2, 0)), X = M = 0, _.mode = 3;
          case 3:
            for (; X < 32; ) {
              if (U === 0) break e;
              U--, M += K[P++] << X, X += 8;
            }
            _.head && (_.head.time = M), 512 & _.flags && (z[0] = 255 & M, z[1] = M >>> 8 & 255, z[2] = M >>> 16 & 255, z[3] = M >>> 24 & 255, _.check = o(_.check, z, 4, 0)), X = M = 0, _.mode = 4;
          case 4:
            for (; X < 16; ) {
              if (U === 0) break e;
              U--, M += K[P++] << X, X += 8;
            }
            _.head && (_.head.xflags = 255 & M, _.head.os = M >> 8), 512 & _.flags && (z[0] = 255 & M, z[1] = M >>> 8 & 255, _.check = o(_.check, z, 2, 0)), X = M = 0, _.mode = 5;
          case 5:
            if (1024 & _.flags) {
              for (; X < 16; ) {
                if (U === 0) break e;
                U--, M += K[P++] << X, X += 8;
              }
              _.length = M, _.head && (_.head.extra_len = M), 512 & _.flags && (z[0] = 255 & M, z[1] = M >>> 8 & 255, _.check = o(_.check, z, 2, 0)), X = M = 0;
            } else _.head && (_.head.extra = null);
            _.mode = 6;
          case 6:
            if (1024 & _.flags && (U < (ne = _.length) && (ne = U), ne && (_.head && (w = _.head.extra_len - _.length, _.head.extra || (_.head.extra = new Array(_.head.extra_len)), a.arraySet(_.head.extra, K, P, ne, w)), 512 & _.flags && (_.check = o(_.check, K, ne, P)), U -= ne, P += ne, _.length -= ne), _.length)) break e;
            _.length = 0, _.mode = 7;
          case 7:
            if (2048 & _.flags) {
              if (U === 0) break e;
              for (ne = 0; w = K[P + ne++], _.head && w && _.length < 65536 && (_.head.name += String.fromCharCode(w)), w && ne < U; ) ;
              if (512 & _.flags && (_.check = o(_.check, K, ne, P)), U -= ne, P += ne, w) break e;
            } else _.head && (_.head.name = null);
            _.length = 0, _.mode = 8;
          case 8:
            if (4096 & _.flags) {
              if (U === 0) break e;
              for (ne = 0; w = K[P + ne++], _.head && w && _.length < 65536 && (_.head.comment += String.fromCharCode(w)), w && ne < U; ) ;
              if (512 & _.flags && (_.check = o(_.check, K, ne, P)), U -= ne, P += ne, w) break e;
            } else _.head && (_.head.comment = null);
            _.mode = 9;
          case 9:
            if (512 & _.flags) {
              for (; X < 16; ) {
                if (U === 0) break e;
                U--, M += K[P++] << X, X += 8;
              }
              if (M !== (65535 & _.check)) {
                T.msg = "header crc mismatch", _.mode = 30;
                break;
              }
              X = M = 0;
            }
            _.head && (_.head.hcrc = _.flags >> 9 & 1, _.head.done = !0), T.adler = _.check = 0, _.mode = 12;
            break;
          case 10:
            for (; X < 32; ) {
              if (U === 0) break e;
              U--, M += K[P++] << X, X += 8;
            }
            T.adler = _.check = m(M), X = M = 0, _.mode = 11;
          case 11:
            if (_.havedict === 0) return T.next_out = G, T.avail_out = Q, T.next_in = P, T.avail_in = U, _.hold = M, _.bits = X, 2;
            T.adler = _.check = 1, _.mode = 12;
          case 12:
            if (j === 5 || j === 6) break e;
          case 13:
            if (_.last) {
              M >>>= 7 & X, X -= 7 & X, _.mode = 27;
              break;
            }
            for (; X < 3; ) {
              if (U === 0) break e;
              U--, M += K[P++] << X, X += 8;
            }
            switch (_.last = 1 & M, X -= 1, 3 & (M >>>= 1)) {
              case 0:
                _.mode = 14;
                break;
              case 1:
                if (B(_), _.mode = 20, j !== 6) break;
                M >>>= 2, X -= 2;
                break e;
              case 2:
                _.mode = 17;
                break;
              case 3:
                T.msg = "invalid block type", _.mode = 30;
            }
            M >>>= 2, X -= 2;
            break;
          case 14:
            for (M >>>= 7 & X, X -= 7 & X; X < 32; ) {
              if (U === 0) break e;
              U--, M += K[P++] << X, X += 8;
            }
            if ((65535 & M) != (M >>> 16 ^ 65535)) {
              T.msg = "invalid stored block lengths", _.mode = 30;
              break;
            }
            if (_.length = 65535 & M, X = M = 0, _.mode = 15, j === 6) break e;
          case 15:
            _.mode = 16;
          case 16:
            if (ne = _.length) {
              if (U < ne && (ne = U), Q < ne && (ne = Q), ne === 0) break e;
              a.arraySet(se, K, P, ne, G), U -= ne, P += ne, Q -= ne, G += ne, _.length -= ne;
              break;
            }
            _.mode = 12;
            break;
          case 17:
            for (; X < 14; ) {
              if (U === 0) break e;
              U--, M += K[P++] << X, X += 8;
            }
            if (_.nlen = 257 + (31 & M), M >>>= 5, X -= 5, _.ndist = 1 + (31 & M), M >>>= 5, X -= 5, _.ncode = 4 + (15 & M), M >>>= 4, X -= 4, 286 < _.nlen || 30 < _.ndist) {
              T.msg = "too many length or distance symbols", _.mode = 30;
              break;
            }
            _.have = 0, _.mode = 18;
          case 18:
            for (; _.have < _.ncode; ) {
              for (; X < 3; ) {
                if (U === 0) break e;
                U--, M += K[P++] << X, X += 8;
              }
              _.lens[W[_.have++]] = 7 & M, M >>>= 3, X -= 3;
            }
            for (; _.have < 19; ) _.lens[W[_.have++]] = 0;
            if (_.lencode = _.lendyn, _.lenbits = 7, ee = { bits: _.lenbits }, ae = l(0, _.lens, 0, 19, _.lencode, 0, _.work, ee), _.lenbits = ee.bits, ae) {
              T.msg = "invalid code lengths set", _.mode = 30;
              break;
            }
            _.have = 0, _.mode = 19;
          case 19:
            for (; _.have < _.nlen + _.ndist; ) {
              for (; Se = (F = _.lencode[M & (1 << _.lenbits) - 1]) >>> 16 & 255, ve = 65535 & F, !((_e = F >>> 24) <= X); ) {
                if (U === 0) break e;
                U--, M += K[P++] << X, X += 8;
              }
              if (ve < 16) M >>>= _e, X -= _e, _.lens[_.have++] = ve;
              else {
                if (ve === 16) {
                  for (D = _e + 2; X < D; ) {
                    if (U === 0) break e;
                    U--, M += K[P++] << X, X += 8;
                  }
                  if (M >>>= _e, X -= _e, _.have === 0) {
                    T.msg = "invalid bit length repeat", _.mode = 30;
                    break;
                  }
                  w = _.lens[_.have - 1], ne = 3 + (3 & M), M >>>= 2, X -= 2;
                } else if (ve === 17) {
                  for (D = _e + 3; X < D; ) {
                    if (U === 0) break e;
                    U--, M += K[P++] << X, X += 8;
                  }
                  X -= _e, w = 0, ne = 3 + (7 & (M >>>= _e)), M >>>= 3, X -= 3;
                } else {
                  for (D = _e + 7; X < D; ) {
                    if (U === 0) break e;
                    U--, M += K[P++] << X, X += 8;
                  }
                  X -= _e, w = 0, ne = 11 + (127 & (M >>>= _e)), M >>>= 7, X -= 7;
                }
                if (_.have + ne > _.nlen + _.ndist) {
                  T.msg = "invalid bit length repeat", _.mode = 30;
                  break;
                }
                for (; ne--; ) _.lens[_.have++] = w;
              }
            }
            if (_.mode === 30) break;
            if (_.lens[256] === 0) {
              T.msg = "invalid code -- missing end-of-block", _.mode = 30;
              break;
            }
            if (_.lenbits = 9, ee = { bits: _.lenbits }, ae = l(f, _.lens, 0, _.nlen, _.lencode, 0, _.work, ee), _.lenbits = ee.bits, ae) {
              T.msg = "invalid literal/lengths set", _.mode = 30;
              break;
            }
            if (_.distbits = 6, _.distcode = _.distdyn, ee = { bits: _.distbits }, ae = l(d, _.lens, _.nlen, _.ndist, _.distcode, 0, _.work, ee), _.distbits = ee.bits, ae) {
              T.msg = "invalid distances set", _.mode = 30;
              break;
            }
            if (_.mode = 20, j === 6) break e;
          case 20:
            _.mode = 21;
          case 21:
            if (6 <= U && 258 <= Q) {
              T.next_out = G, T.avail_out = Q, T.next_in = P, T.avail_in = U, _.hold = M, _.bits = X, h(T, le), G = T.next_out, se = T.output, Q = T.avail_out, P = T.next_in, K = T.input, U = T.avail_in, M = _.hold, X = _.bits, _.mode === 12 && (_.back = -1);
              break;
            }
            for (_.back = 0; Se = (F = _.lencode[M & (1 << _.lenbits) - 1]) >>> 16 & 255, ve = 65535 & F, !((_e = F >>> 24) <= X); ) {
              if (U === 0) break e;
              U--, M += K[P++] << X, X += 8;
            }
            if (Se && !(240 & Se)) {
              for (Re = _e, Ue = Se, He = ve; Se = (F = _.lencode[He + ((M & (1 << Re + Ue) - 1) >> Re)]) >>> 16 & 255, ve = 65535 & F, !(Re + (_e = F >>> 24) <= X); ) {
                if (U === 0) break e;
                U--, M += K[P++] << X, X += 8;
              }
              M >>>= Re, X -= Re, _.back += Re;
            }
            if (M >>>= _e, X -= _e, _.back += _e, _.length = ve, Se === 0) {
              _.mode = 26;
              break;
            }
            if (32 & Se) {
              _.back = -1, _.mode = 12;
              break;
            }
            if (64 & Se) {
              T.msg = "invalid literal/length code", _.mode = 30;
              break;
            }
            _.extra = 15 & Se, _.mode = 22;
          case 22:
            if (_.extra) {
              for (D = _.extra; X < D; ) {
                if (U === 0) break e;
                U--, M += K[P++] << X, X += 8;
              }
              _.length += M & (1 << _.extra) - 1, M >>>= _.extra, X -= _.extra, _.back += _.extra;
            }
            _.was = _.length, _.mode = 23;
          case 23:
            for (; Se = (F = _.distcode[M & (1 << _.distbits) - 1]) >>> 16 & 255, ve = 65535 & F, !((_e = F >>> 24) <= X); ) {
              if (U === 0) break e;
              U--, M += K[P++] << X, X += 8;
            }
            if (!(240 & Se)) {
              for (Re = _e, Ue = Se, He = ve; Se = (F = _.distcode[He + ((M & (1 << Re + Ue) - 1) >> Re)]) >>> 16 & 255, ve = 65535 & F, !(Re + (_e = F >>> 24) <= X); ) {
                if (U === 0) break e;
                U--, M += K[P++] << X, X += 8;
              }
              M >>>= Re, X -= Re, _.back += Re;
            }
            if (M >>>= _e, X -= _e, _.back += _e, 64 & Se) {
              T.msg = "invalid distance code", _.mode = 30;
              break;
            }
            _.offset = ve, _.extra = 15 & Se, _.mode = 24;
          case 24:
            if (_.extra) {
              for (D = _.extra; X < D; ) {
                if (U === 0) break e;
                U--, M += K[P++] << X, X += 8;
              }
              _.offset += M & (1 << _.extra) - 1, M >>>= _.extra, X -= _.extra, _.back += _.extra;
            }
            if (_.offset > _.dmax) {
              T.msg = "invalid distance too far back", _.mode = 30;
              break;
            }
            _.mode = 25;
          case 25:
            if (Q === 0) break e;
            if (ne = le - Q, _.offset > ne) {
              if ((ne = _.offset - ne) > _.whave && _.sane) {
                T.msg = "invalid distance too far back", _.mode = 30;
                break;
              }
              ye = ne > _.wnext ? (ne -= _.wnext, _.wsize - ne) : _.wnext - ne, ne > _.length && (ne = _.length), Te = _.window;
            } else Te = se, ye = G - _.offset, ne = _.length;
            for (Q < ne && (ne = Q), Q -= ne, _.length -= ne; se[G++] = Te[ye++], --ne; ) ;
            _.length === 0 && (_.mode = 21);
            break;
          case 26:
            if (Q === 0) break e;
            se[G++] = _.length, Q--, _.mode = 21;
            break;
          case 27:
            if (_.wrap) {
              for (; X < 32; ) {
                if (U === 0) break e;
                U--, M |= K[P++] << X, X += 8;
              }
              if (le -= Q, T.total_out += le, _.total += le, le && (T.adler = _.check = _.flags ? o(_.check, se, le, G - le) : s(_.check, se, le, G - le)), le = Q, (_.flags ? M : m(M)) !== _.check) {
                T.msg = "incorrect data check", _.mode = 30;
                break;
              }
              X = M = 0;
            }
            _.mode = 28;
          case 28:
            if (_.wrap && _.flags) {
              for (; X < 32; ) {
                if (U === 0) break e;
                U--, M += K[P++] << X, X += 8;
              }
              if (M !== (4294967295 & _.total)) {
                T.msg = "incorrect length check", _.mode = 30;
                break;
              }
              X = M = 0;
            }
            _.mode = 29;
          case 29:
            ae = 1;
            break e;
          case 30:
            ae = -3;
            break e;
          case 31:
            return -4;
          case 32:
          default:
            return c;
        }
        return T.next_out = G, T.avail_out = Q, T.next_in = P, T.avail_in = U, _.hold = M, _.bits = X, (_.wsize || le !== T.avail_out && _.mode < 30 && (_.mode < 27 || j !== 4)) && Z(T, T.output, T.next_out, le - T.avail_out) ? (_.mode = 31, -4) : (de -= T.avail_in, le -= T.avail_out, T.total_in += de, T.total_out += le, _.total += le, _.wrap && le && (T.adler = _.check = _.flags ? o(_.check, se, le, T.next_out - le) : s(_.check, se, le, T.next_out - le)), T.data_type = _.bits + (_.last ? 64 : 0) + (_.mode === 12 ? 128 : 0) + (_.mode === 20 || _.mode === 15 ? 256 : 0), (de == 0 && le === 0 || j === 4) && ae === u && (ae = -5), ae);
      }, i.inflateEnd = function(T) {
        if (!T || !T.state) return c;
        var j = T.state;
        return j.window && (j.window = null), T.state = null, u;
      }, i.inflateGetHeader = function(T, j) {
        var _;
        return T && T.state && 2 & (_ = T.state).wrap ? ((_.head = j).done = !1, u) : c;
      }, i.inflateSetDictionary = function(T, j) {
        var _, K = j.length;
        return T && T.state ? (_ = T.state).wrap !== 0 && _.mode !== 11 ? c : _.mode === 11 && s(1, j, K, 0) !== _.check ? -3 : Z(T, j, K, K) ? (_.mode = 31, -4) : (_.havedict = 1, u) : c;
      }, i.inflateInfo = "pako inflate (from Nodeca project)";
    }, { "../utils/common": 41, "./adler32": 43, "./crc32": 45, "./inffast": 48, "./inftrees": 50 }], 50: [function(e, n, i) {
      var a = e("../utils/common"), s = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258, 0, 0], o = [16, 16, 16, 16, 16, 16, 16, 16, 17, 17, 17, 17, 18, 18, 18, 18, 19, 19, 19, 19, 20, 20, 20, 20, 21, 21, 21, 21, 16, 72, 78], h = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577, 0, 0], l = [16, 16, 16, 16, 17, 17, 18, 18, 19, 19, 20, 20, 21, 21, 22, 22, 23, 23, 24, 24, 25, 25, 26, 26, 27, 27, 28, 28, 29, 29, 64, 64];
      n.exports = function(f, d, u, c, g, p, y, m) {
        var A, E, x, S, R, v, O, C, B, Z = m.bits, T = 0, j = 0, _ = 0, K = 0, se = 0, P = 0, G = 0, U = 0, Q = 0, M = 0, X = null, de = 0, le = new a.Buf16(16), ne = new a.Buf16(16), ye = null, Te = 0;
        for (T = 0; T <= 15; T++) le[T] = 0;
        for (j = 0; j < c; j++) le[d[u + j]]++;
        for (se = Z, K = 15; 1 <= K && le[K] === 0; K--) ;
        if (K < se && (se = K), K === 0) return g[p++] = 20971520, g[p++] = 20971520, m.bits = 1, 0;
        for (_ = 1; _ < K && le[_] === 0; _++) ;
        for (se < _ && (se = _), T = U = 1; T <= 15; T++) if (U <<= 1, (U -= le[T]) < 0) return -1;
        if (0 < U && (f === 0 || K !== 1)) return -1;
        for (ne[1] = 0, T = 1; T < 15; T++) ne[T + 1] = ne[T] + le[T];
        for (j = 0; j < c; j++) d[u + j] !== 0 && (y[ne[d[u + j]]++] = j);
        if (v = f === 0 ? (X = ye = y, 19) : f === 1 ? (X = s, de -= 257, ye = o, Te -= 257, 256) : (X = h, ye = l, -1), T = _, R = p, G = j = M = 0, x = -1, S = (Q = 1 << (P = se)) - 1, f === 1 && 852 < Q || f === 2 && 592 < Q) return 1;
        for (; ; ) {
          for (O = T - G, B = y[j] < v ? (C = 0, y[j]) : y[j] > v ? (C = ye[Te + y[j]], X[de + y[j]]) : (C = 96, 0), A = 1 << T - G, _ = E = 1 << P; g[R + (M >> G) + (E -= A)] = O << 24 | C << 16 | B | 0, E !== 0; ) ;
          for (A = 1 << T - 1; M & A; ) A >>= 1;
          if (A !== 0 ? (M &= A - 1, M += A) : M = 0, j++, --le[T] == 0) {
            if (T === K) break;
            T = d[u + y[j]];
          }
          if (se < T && (M & S) !== x) {
            for (G === 0 && (G = se), R += _, U = 1 << (P = T - G); P + G < K && !((U -= le[P + G]) <= 0); ) P++, U <<= 1;
            if (Q += 1 << P, f === 1 && 852 < Q || f === 2 && 592 < Q) return 1;
            g[x = M & S] = se << 24 | P << 16 | R - p | 0;
          }
        }
        return M !== 0 && (g[R + M] = T - G << 24 | 64 << 16 | 0), m.bits = se, 0;
      };
    }, { "../utils/common": 41 }], 51: [function(e, n, i) {
      n.exports = { 2: "need dictionary", 1: "stream end", 0: "", "-1": "file error", "-2": "stream error", "-3": "data error", "-4": "insufficient memory", "-5": "buffer error", "-6": "incompatible version" };
    }, {}], 52: [function(e, n, i) {
      var a = e("../utils/common"), s = 0, o = 1;
      function h(F) {
        for (var z = F.length; 0 <= --z; ) F[z] = 0;
      }
      var l = 0, f = 29, d = 256, u = d + 1 + f, c = 30, g = 19, p = 2 * u + 1, y = 15, m = 16, A = 7, E = 256, x = 16, S = 17, R = 18, v = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0], O = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13], C = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 3, 7], B = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15], Z = new Array(2 * (u + 2));
      h(Z);
      var T = new Array(2 * c);
      h(T);
      var j = new Array(512);
      h(j);
      var _ = new Array(256);
      h(_);
      var K = new Array(f);
      h(K);
      var se, P, G, U = new Array(c);
      function Q(F, z, W, ie, $) {
        this.static_tree = F, this.extra_bits = z, this.extra_base = W, this.elems = ie, this.max_length = $, this.has_stree = F && F.length;
      }
      function M(F, z) {
        this.dyn_tree = F, this.max_code = 0, this.stat_desc = z;
      }
      function X(F) {
        return F < 256 ? j[F] : j[256 + (F >>> 7)];
      }
      function de(F, z) {
        F.pending_buf[F.pending++] = 255 & z, F.pending_buf[F.pending++] = z >>> 8 & 255;
      }
      function le(F, z, W) {
        F.bi_valid > m - W ? (F.bi_buf |= z << F.bi_valid & 65535, de(F, F.bi_buf), F.bi_buf = z >> m - F.bi_valid, F.bi_valid += W - m) : (F.bi_buf |= z << F.bi_valid & 65535, F.bi_valid += W);
      }
      function ne(F, z, W) {
        le(F, W[2 * z], W[2 * z + 1]);
      }
      function ye(F, z) {
        for (var W = 0; W |= 1 & F, F >>>= 1, W <<= 1, 0 < --z; ) ;
        return W >>> 1;
      }
      function Te(F, z, W) {
        var ie, $, re = new Array(y + 1), he = 0;
        for (ie = 1; ie <= y; ie++) re[ie] = he = he + W[ie - 1] << 1;
        for ($ = 0; $ <= z; $++) {
          var oe = F[2 * $ + 1];
          oe !== 0 && (F[2 * $] = ye(re[oe]++, oe));
        }
      }
      function _e(F) {
        var z;
        for (z = 0; z < u; z++) F.dyn_ltree[2 * z] = 0;
        for (z = 0; z < c; z++) F.dyn_dtree[2 * z] = 0;
        for (z = 0; z < g; z++) F.bl_tree[2 * z] = 0;
        F.dyn_ltree[2 * E] = 1, F.opt_len = F.static_len = 0, F.last_lit = F.matches = 0;
      }
      function Se(F) {
        8 < F.bi_valid ? de(F, F.bi_buf) : 0 < F.bi_valid && (F.pending_buf[F.pending++] = F.bi_buf), F.bi_buf = 0, F.bi_valid = 0;
      }
      function ve(F, z, W, ie) {
        var $ = 2 * z, re = 2 * W;
        return F[$] < F[re] || F[$] === F[re] && ie[z] <= ie[W];
      }
      function Re(F, z, W) {
        for (var ie = F.heap[W], $ = W << 1; $ <= F.heap_len && ($ < F.heap_len && ve(z, F.heap[$ + 1], F.heap[$], F.depth) && $++, !ve(z, ie, F.heap[$], F.depth)); ) F.heap[W] = F.heap[$], W = $, $ <<= 1;
        F.heap[W] = ie;
      }
      function Ue(F, z, W) {
        var ie, $, re, he, oe = 0;
        if (F.last_lit !== 0) for (; ie = F.pending_buf[F.d_buf + 2 * oe] << 8 | F.pending_buf[F.d_buf + 2 * oe + 1], $ = F.pending_buf[F.l_buf + oe], oe++, ie === 0 ? ne(F, $, z) : (ne(F, (re = _[$]) + d + 1, z), (he = v[re]) !== 0 && le(F, $ -= K[re], he), ne(F, re = X(--ie), W), (he = O[re]) !== 0 && le(F, ie -= U[re], he)), oe < F.last_lit; ) ;
        ne(F, E, z);
      }
      function He(F, z) {
        var W, ie, $, re = z.dyn_tree, he = z.stat_desc.static_tree, oe = z.stat_desc.has_stree, ge = z.stat_desc.elems, Ie = -1;
        for (F.heap_len = 0, F.heap_max = p, W = 0; W < ge; W++) re[2 * W] !== 0 ? (F.heap[++F.heap_len] = Ie = W, F.depth[W] = 0) : re[2 * W + 1] = 0;
        for (; F.heap_len < 2; ) re[2 * ($ = F.heap[++F.heap_len] = Ie < 2 ? ++Ie : 0)] = 1, F.depth[$] = 0, F.opt_len--, oe && (F.static_len -= he[2 * $ + 1]);
        for (z.max_code = Ie, W = F.heap_len >> 1; 1 <= W; W--) Re(F, re, W);
        for ($ = ge; W = F.heap[1], F.heap[1] = F.heap[F.heap_len--], Re(F, re, 1), ie = F.heap[1], F.heap[--F.heap_max] = W, F.heap[--F.heap_max] = ie, re[2 * $] = re[2 * W] + re[2 * ie], F.depth[$] = (F.depth[W] >= F.depth[ie] ? F.depth[W] : F.depth[ie]) + 1, re[2 * W + 1] = re[2 * ie + 1] = $, F.heap[1] = $++, Re(F, re, 1), 2 <= F.heap_len; ) ;
        F.heap[--F.heap_max] = F.heap[1], function(be, De) {
          var Je, Xe, Ve, Fe, qe, Bt, Ye = De.dyn_tree, cr = De.max_code, Pr = De.stat_desc.static_tree, Tt = De.stat_desc.has_stree, Et = De.stat_desc.extra_bits, fr = De.stat_desc.extra_base, ke = De.stat_desc.max_length, ze = 0;
          for (Fe = 0; Fe <= y; Fe++) be.bl_count[Fe] = 0;
          for (Ye[2 * be.heap[be.heap_max] + 1] = 0, Je = be.heap_max + 1; Je < p; Je++) ke < (Fe = Ye[2 * Ye[2 * (Xe = be.heap[Je]) + 1] + 1] + 1) && (Fe = ke, ze++), Ye[2 * Xe + 1] = Fe, cr < Xe || (be.bl_count[Fe]++, qe = 0, fr <= Xe && (qe = Et[Xe - fr]), Bt = Ye[2 * Xe], be.opt_len += Bt * (Fe + qe), Tt && (be.static_len += Bt * (Pr[2 * Xe + 1] + qe)));
          if (ze !== 0) {
            do {
              for (Fe = ke - 1; be.bl_count[Fe] === 0; ) Fe--;
              be.bl_count[Fe]--, be.bl_count[Fe + 1] += 2, be.bl_count[ke]--, ze -= 2;
            } while (0 < ze);
            for (Fe = ke; Fe !== 0; Fe--) for (Xe = be.bl_count[Fe]; Xe !== 0; ) cr < (Ve = be.heap[--Je]) || (Ye[2 * Ve + 1] !== Fe && (be.opt_len += (Fe - Ye[2 * Ve + 1]) * Ye[2 * Ve], Ye[2 * Ve + 1] = Fe), Xe--);
          }
        }(F, z), Te(re, Ie, F.bl_count);
      }
      function w(F, z, W) {
        var ie, $, re = -1, he = z[1], oe = 0, ge = 7, Ie = 4;
        for (he === 0 && (ge = 138, Ie = 3), z[2 * (W + 1) + 1] = 65535, ie = 0; ie <= W; ie++) $ = he, he = z[2 * (ie + 1) + 1], ++oe < ge && $ === he || (oe < Ie ? F.bl_tree[2 * $] += oe : $ !== 0 ? ($ !== re && F.bl_tree[2 * $]++, F.bl_tree[2 * x]++) : oe <= 10 ? F.bl_tree[2 * S]++ : F.bl_tree[2 * R]++, re = $, Ie = (oe = 0) === he ? (ge = 138, 3) : $ === he ? (ge = 6, 3) : (ge = 7, 4));
      }
      function ae(F, z, W) {
        var ie, $, re = -1, he = z[1], oe = 0, ge = 7, Ie = 4;
        for (he === 0 && (ge = 138, Ie = 3), ie = 0; ie <= W; ie++) if ($ = he, he = z[2 * (ie + 1) + 1], !(++oe < ge && $ === he)) {
          if (oe < Ie) for (; ne(F, $, F.bl_tree), --oe != 0; ) ;
          else $ !== 0 ? ($ !== re && (ne(F, $, F.bl_tree), oe--), ne(F, x, F.bl_tree), le(F, oe - 3, 2)) : oe <= 10 ? (ne(F, S, F.bl_tree), le(F, oe - 3, 3)) : (ne(F, R, F.bl_tree), le(F, oe - 11, 7));
          re = $, Ie = (oe = 0) === he ? (ge = 138, 3) : $ === he ? (ge = 6, 3) : (ge = 7, 4);
        }
      }
      h(U);
      var ee = !1;
      function D(F, z, W, ie) {
        le(F, (l << 1) + (ie ? 1 : 0), 3), function($, re, he, oe) {
          Se($), de($, he), de($, ~he), a.arraySet($.pending_buf, $.window, re, he, $.pending), $.pending += he;
        }(F, z, W);
      }
      i._tr_init = function(F) {
        ee || (function() {
          var z, W, ie, $, re, he = new Array(y + 1);
          for ($ = ie = 0; $ < f - 1; $++) for (K[$] = ie, z = 0; z < 1 << v[$]; z++) _[ie++] = $;
          for (_[ie - 1] = $, $ = re = 0; $ < 16; $++) for (U[$] = re, z = 0; z < 1 << O[$]; z++) j[re++] = $;
          for (re >>= 7; $ < c; $++) for (U[$] = re << 7, z = 0; z < 1 << O[$] - 7; z++) j[256 + re++] = $;
          for (W = 0; W <= y; W++) he[W] = 0;
          for (z = 0; z <= 143; ) Z[2 * z + 1] = 8, z++, he[8]++;
          for (; z <= 255; ) Z[2 * z + 1] = 9, z++, he[9]++;
          for (; z <= 279; ) Z[2 * z + 1] = 7, z++, he[7]++;
          for (; z <= 287; ) Z[2 * z + 1] = 8, z++, he[8]++;
          for (Te(Z, u + 1, he), z = 0; z < c; z++) T[2 * z + 1] = 5, T[2 * z] = ye(z, 5);
          se = new Q(Z, v, d + 1, u, y), P = new Q(T, O, 0, c, y), G = new Q(new Array(0), C, 0, g, A);
        }(), ee = !0), F.l_desc = new M(F.dyn_ltree, se), F.d_desc = new M(F.dyn_dtree, P), F.bl_desc = new M(F.bl_tree, G), F.bi_buf = 0, F.bi_valid = 0, _e(F);
      }, i._tr_stored_block = D, i._tr_flush_block = function(F, z, W, ie) {
        var $, re, he = 0;
        0 < F.level ? (F.strm.data_type === 2 && (F.strm.data_type = function(oe) {
          var ge, Ie = 4093624447;
          for (ge = 0; ge <= 31; ge++, Ie >>>= 1) if (1 & Ie && oe.dyn_ltree[2 * ge] !== 0) return s;
          if (oe.dyn_ltree[18] !== 0 || oe.dyn_ltree[20] !== 0 || oe.dyn_ltree[26] !== 0) return o;
          for (ge = 32; ge < d; ge++) if (oe.dyn_ltree[2 * ge] !== 0) return o;
          return s;
        }(F)), He(F, F.l_desc), He(F, F.d_desc), he = function(oe) {
          var ge;
          for (w(oe, oe.dyn_ltree, oe.l_desc.max_code), w(oe, oe.dyn_dtree, oe.d_desc.max_code), He(oe, oe.bl_desc), ge = g - 1; 3 <= ge && oe.bl_tree[2 * B[ge] + 1] === 0; ge--) ;
          return oe.opt_len += 3 * (ge + 1) + 5 + 5 + 4, ge;
        }(F), $ = F.opt_len + 3 + 7 >>> 3, (re = F.static_len + 3 + 7 >>> 3) <= $ && ($ = re)) : $ = re = W + 5, W + 4 <= $ && z !== -1 ? D(F, z, W, ie) : F.strategy === 4 || re === $ ? (le(F, 2 + (ie ? 1 : 0), 3), Ue(F, Z, T)) : (le(F, 4 + (ie ? 1 : 0), 3), function(oe, ge, Ie, be) {
          var De;
          for (le(oe, ge - 257, 5), le(oe, Ie - 1, 5), le(oe, be - 4, 4), De = 0; De < be; De++) le(oe, oe.bl_tree[2 * B[De] + 1], 3);
          ae(oe, oe.dyn_ltree, ge - 1), ae(oe, oe.dyn_dtree, Ie - 1);
        }(F, F.l_desc.max_code + 1, F.d_desc.max_code + 1, he + 1), Ue(F, F.dyn_ltree, F.dyn_dtree)), _e(F), ie && Se(F);
      }, i._tr_tally = function(F, z, W) {
        return F.pending_buf[F.d_buf + 2 * F.last_lit] = z >>> 8 & 255, F.pending_buf[F.d_buf + 2 * F.last_lit + 1] = 255 & z, F.pending_buf[F.l_buf + F.last_lit] = 255 & W, F.last_lit++, z === 0 ? F.dyn_ltree[2 * W]++ : (F.matches++, z--, F.dyn_ltree[2 * (_[W] + d + 1)]++, F.dyn_dtree[2 * X(z)]++), F.last_lit === F.lit_bufsize - 1;
      }, i._tr_align = function(F) {
        le(F, 2, 3), ne(F, E, Z), function(z) {
          z.bi_valid === 16 ? (de(z, z.bi_buf), z.bi_buf = 0, z.bi_valid = 0) : 8 <= z.bi_valid && (z.pending_buf[z.pending++] = 255 & z.bi_buf, z.bi_buf >>= 8, z.bi_valid -= 8);
        }(F);
      };
    }, { "../utils/common": 41 }], 53: [function(e, n, i) {
      n.exports = function() {
        this.input = null, this.next_in = 0, this.avail_in = 0, this.total_in = 0, this.output = null, this.next_out = 0, this.avail_out = 0, this.total_out = 0, this.msg = "", this.state = null, this.data_type = 2, this.adler = 0;
      };
    }, {}], 54: [function(e, n, i) {
      (function(a) {
        (function(s, o) {
          if (!s.setImmediate) {
            var h, l, f, d, u = 1, c = {}, g = !1, p = s.document, y = Object.getPrototypeOf && Object.getPrototypeOf(s);
            y = y && y.setTimeout ? y : s, h = {}.toString.call(s.process) === "[object process]" ? function(x) {
              process.nextTick(function() {
                A(x);
              });
            } : function() {
              if (s.postMessage && !s.importScripts) {
                var x = !0, S = s.onmessage;
                return s.onmessage = function() {
                  x = !1;
                }, s.postMessage("", "*"), s.onmessage = S, x;
              }
            }() ? (d = "setImmediate$" + Math.random() + "$", s.addEventListener ? s.addEventListener("message", E, !1) : s.attachEvent("onmessage", E), function(x) {
              s.postMessage(d + x, "*");
            }) : s.MessageChannel ? ((f = new MessageChannel()).port1.onmessage = function(x) {
              A(x.data);
            }, function(x) {
              f.port2.postMessage(x);
            }) : p && "onreadystatechange" in p.createElement("script") ? (l = p.documentElement, function(x) {
              var S = p.createElement("script");
              S.onreadystatechange = function() {
                A(x), S.onreadystatechange = null, l.removeChild(S), S = null;
              }, l.appendChild(S);
            }) : function(x) {
              setTimeout(A, 0, x);
            }, y.setImmediate = function(x) {
              typeof x != "function" && (x = new Function("" + x));
              for (var S = new Array(arguments.length - 1), R = 0; R < S.length; R++) S[R] = arguments[R + 1];
              var v = { callback: x, args: S };
              return c[u] = v, h(u), u++;
            }, y.clearImmediate = m;
          }
          function m(x) {
            delete c[x];
          }
          function A(x) {
            if (g) setTimeout(A, 0, x);
            else {
              var S = c[x];
              if (S) {
                g = !0;
                try {
                  (function(R) {
                    var v = R.callback, O = R.args;
                    switch (O.length) {
                      case 0:
                        v();
                        break;
                      case 1:
                        v(O[0]);
                        break;
                      case 2:
                        v(O[0], O[1]);
                        break;
                      case 3:
                        v(O[0], O[1], O[2]);
                        break;
                      default:
                        v.apply(o, O);
                    }
                  })(S);
                } finally {
                  m(x), g = !1;
                }
              }
            }
          }
          function E(x) {
            x.source === s && typeof x.data == "string" && x.data.indexOf(d) === 0 && A(+x.data.slice(d.length));
          }
        })(typeof self > "u" ? a === void 0 ? this : a : self);
      }).call(this, typeof mr < "u" ? mr : typeof self < "u" ? self : typeof window < "u" ? window : {});
    }, {}] }, {}, [10])(10);
  });
})(ma);
var Ti = ma.exports;
const yt = /* @__PURE__ */ bi(Ti);
class Ei {
  constructor() {
    ce(this, "ns", {
      w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      a: "http://schemas.openxmlformats.org/drawingml/2006/main",
      pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
      wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    });
    ce(this, "hwpxNs", {
      ha: "http://www.hancom.co.kr/hwpml/2011/app",
      hp: "http://www.hancom.co.kr/hwpml/2011/paragraph",
      hp10: "http://www.hancom.co.kr/hwpml/2016/paragraph",
      hs: "http://www.hancom.co.kr/hwpml/2011/section",
      hc: "http://www.hancom.co.kr/hwpml/2011/core",
      hh: "http://www.hancom.co.kr/hwpml/2011/head",
      hhs: "http://www.hancom.co.kr/hwpml/2011/history",
      hm: "http://www.hancom.co.kr/hwpml/2011/master-page",
      hpf: "http://www.hancom.co.kr/schema/2011/hpf",
      dc: "http://purl.org/dc/elements/1.1/",
      opf: "http://www.idpf.org/2007/opf/",
      ooxmlchart: "http://www.hancom.co.kr/hwpml/2016/ooxmlchart",
      epub: "http://www.idpf.org/2007/ops",
      config: "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
      hv: "http://www.hancom.co.kr/hwpml/2011/version",
      ocf: "urn:oasis:names:tc:opendocument:xmlns:container",
      odf: "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
      rdf: "http://www.w3.org/1999/02/22-rdf-syntax-ns#",
      odfPkg: "http://www.hancom.co.kr/hwpml/2016/meta/pkg#Document"
    });
    ce(this, "HWPUNIT_PER_INCH", 7200);
    ce(this, "EMU_PER_INCH", 914400);
    // 페이지 설정 (기본값 A4)
    ce(this, "PAGE_WIDTH", 59530);
    ce(this, "PAGE_HEIGHT", 84190);
    ce(this, "MARGIN_TOP", 5668);
    ce(this, "MARGIN_LEFT", 8505);
    ce(this, "MARGIN_RIGHT", 8505);
    ce(this, "MARGIN_BOTTOM", 4252);
    ce(this, "HEADER_MARGIN", 4252);
    ce(this, "FOOTER_MARGIN", 4252);
    ce(this, "GUTTER_MARGIN", 0);
    ce(this, "COL_COUNT", 1);
    ce(this, "COL_GAP", 0);
    ce(this, "COL_TYPE", "ONE");
    ce(this, "TEXT_WIDTH");
    ce(this, "charShapes", /* @__PURE__ */ new Map());
    ce(this, "paraShapes", /* @__PURE__ */ new Map());
    ce(this, "borderFills", /* @__PURE__ */ new Map());
    ce(this, "idCounter", 3121190098);
    ce(this, "tableIdCounter", 2085132891);
    ce(this, "borderFillIdCounter", 7);
    // IDs 1-6 are pre-initialized
    ce(this, "imageCounter", 1);
    ce(this, "currentVertPos", 0);
    // Footnote/Endnote
    ce(this, "footnoteNumber", 0);
    ce(this, "endnoteNumber", 0);
    ce(this, "metadata", {
      title: "",
      creator: "",
      subject: "",
      description: ""
    });
    ce(this, "relationships", /* @__PURE__ */ new Map());
    ce(this, "images", /* @__PURE__ */ new Map());
    // Fonts
    ce(this, "langFontFaces", {
      HANGUL: /* @__PURE__ */ new Map(),
      LATIN: /* @__PURE__ */ new Map(),
      HANJA: /* @__PURE__ */ new Map(),
      JAPANESE: /* @__PURE__ */ new Map(),
      OTHER: /* @__PURE__ */ new Map(),
      SYMBOL: /* @__PURE__ */ new Map(),
      USER: /* @__PURE__ */ new Map()
    });
    // Numbering
    ce(this, "numberingMap", /* @__PURE__ */ new Map());
    // numId -> abstractNumId
    ce(this, "abstractNumMap", /* @__PURE__ */ new Map());
    // abstractNumId -> levels
    ce(this, "listCounters", /* @__PURE__ */ new Map());
    // "numId:ilvl" -> counter
    ce(this, "docStyles", /* @__PURE__ */ new Map());
    ce(this, "docxStyleToHwpxId", {});
    ce(this, "footnoteMap", /* @__PURE__ */ new Map());
    ce(this, "endnoteMap", /* @__PURE__ */ new Map());
    ce(this, "isFirstParagraph", !0);
    this.TEXT_WIDTH = this.PAGE_WIDTH - this.MARGIN_LEFT - this.MARGIN_RIGHT - this.GUTTER_MARGIN, this.registerFontForLang("HANGUL", "함초롬바탕"), this.registerFontForLang("LATIN", "함초롬바탕"), this.registerFontForLang("HANJA", "함초롬바탕"), this.registerFontForLang("JAPANESE", "함초롬바탕"), this.registerFontForLang("OTHER", "함초롬바탕"), this.registerFontForLang("SYMBOL", "함초롬바탕"), this.registerFontForLang("USER", "함초롬바탕"), this.initDocxStyleToHwpxId(), this.initDefaultStyles();
  }
  registerFontForLang(r, e) {
    const n = this.langFontFaces[r];
    return n ? (n.has(e) || n.set(e, String(n.size)), parseInt(n.get(e))) : 0;
  }
  initDocxStyleToHwpxId() {
    this.docxStyleToHwpxId = {
      Normal: 0,
      Heading1: 12,
      Heading2: 14,
      Heading3: 16,
      Heading4: 18,
      Heading5: 20,
      Heading6: 22,
      Heading7: 24,
      Heading8: 26,
      Heading9: 28,
      Title: 46,
      Subtitle: 41,
      ListParagraph: 35,
      NoSpacing: 36,
      Quote: 38,
      IntenseQuote: 32,
      Header: 10,
      Footer: 7,
      FootnoteText: 51,
      EndnoteText: 49,
      TOCHeading: 45,
      TOC1: 53,
      TOC2: 54,
      TOC3: 55,
      TOC4: 56,
      TOC5: 57,
      TOC6: 58,
      TOC7: 59,
      TOC8: 60,
      TOC9: 61,
      Caption: 2,
      TableofFigures: 52
    };
  }
  initDefaultStyles() {
    this.borderFills.set("1", {
      id: "1",
      leftBorder: { type: "NONE", width: "0.1 mm", color: "#000000" },
      rightBorder: { type: "NONE", width: "0.1 mm", color: "#000000" },
      topBorder: { type: "NONE", width: "0.1 mm", color: "#000000" },
      bottomBorder: { type: "NONE", width: "0.1 mm", color: "#000000" },
      diagonal: null,
      slash: { type: "NONE", width: "0.1 mm", color: "#000000" },
      backSlash: { type: "NONE", width: "0.1 mm", color: "#000000" },
      backColor: "#FFFFFF"
    }), this.borderFills.set("2", {
      id: "2",
      leftBorder: { type: "SOLID", width: "0.1 mm", color: "#000000" },
      rightBorder: { type: "SOLID", width: "0.1 mm", color: "#000000" },
      topBorder: { type: "SOLID", width: "0.1 mm", color: "#000000" },
      bottomBorder: { type: "SOLID", width: "0.1 mm", color: "#000000" },
      diagonal: null,
      slash: { type: "NONE", width: "0.1 mm", color: "#000000" },
      backSlash: { type: "NONE", width: "0.1 mm", color: "#000000" },
      backColor: "#FFFFFF"
    }), this.borderFills.set("3", {
      id: "3",
      leftBorder: { type: "SOLID", width: "0.5 mm", color: "#000000" },
      rightBorder: { type: "SOLID", width: "0.5 mm", color: "#000000" },
      topBorder: { type: "SOLID", width: "0.5 mm", color: "#000000" },
      bottomBorder: { type: "SOLID", width: "0.5 mm", color: "#000000" },
      diagonal: null,
      slash: { type: "NONE", width: "0.1 mm", color: "#000000" },
      backSlash: { type: "NONE", width: "0.1 mm", color: "#000000" },
      backColor: "#FFFFFF"
    }), this.borderFills.set("4", {
      id: "4",
      leftBorder: { type: "SOLID", width: "0.1 mm", color: "#D0D0D0" },
      rightBorder: { type: "SOLID", width: "0.1 mm", color: "#D0D0D0" },
      topBorder: { type: "SOLID", width: "0.1 mm", color: "#D0D0D0" },
      bottomBorder: { type: "SOLID", width: "0.1 mm", color: "#D0D0D0" },
      diagonal: null,
      slash: { type: "NONE", width: "0.1 mm", color: "#000000" },
      backSlash: { type: "NONE", width: "0.1 mm", color: "#000000" },
      backColor: "#FFFFFF"
    }), this.borderFills.set("5", {
      id: "5",
      leftBorder: { type: "SOLID", width: "0.12 mm", color: "#000000" },
      rightBorder: { type: "SOLID", width: "0.12 mm", color: "#000000" },
      topBorder: { type: "SOLID", width: "0.12 mm", color: "#000000" },
      bottomBorder: { type: "SOLID", width: "0.12 mm", color: "#000000" },
      diagonal: null,
      slash: { type: "NONE", width: "0.1 mm", color: "#000000" },
      backSlash: { type: "NONE", width: "0.1 mm", color: "#000000" },
      backColor: "#FFFFFF"
    }), this.borderFills.set("6", {
      id: "6",
      leftBorder: { type: "NONE", width: "0.1 mm", color: "#000000" },
      rightBorder: { type: "NONE", width: "0.1 mm", color: "#000000" },
      topBorder: { type: "NONE", width: "0.1 mm", color: "#000000" },
      bottomBorder: { type: "NONE", width: "0.1 mm", color: "#000000" },
      diagonal: null,
      slash: { type: "NONE", width: "0.1 mm", color: "#000000" },
      backSlash: { type: "NONE", width: "0.1 mm", color: "#000000" },
      backColor: "#000000"
    });
    const r = {
      hangulId: 0,
      latinId: 0,
      hanjaId: 0,
      japaneseId: 0,
      otherId: 0,
      symbolId: 0,
      userId: 0
    }, e = (n, i, a = !1, s = !1) => ({
      id: n,
      height: i,
      textColor: "#000000",
      shadeColor: "none",
      borderFillIDRef: "1",
      bold: a,
      italic: s,
      underline: !1,
      strike: !1,
      supscript: !1,
      subscript: !1,
      fontIds: { ...r },
      fontId: 0
    });
    this.charShapes.set("0", e("0", "1000")), this.charShapes.set("1", e("1", "1000")), this.charShapes.set("2", e("2", "1000")), this.charShapes.set("19", e("19", "1000", !0)), this.charShapes.set("24", e("24", "1400", !0)), this.charShapes.set("28", e("28", "900")), this.paraShapes.set("0", {
      id: "0",
      align: "LEFT",
      heading: "NONE",
      level: "0",
      headingIdRef: "0",
      leftMargin: 0,
      rightMargin: 0,
      indent: 0,
      prevSpacing: 0,
      nextSpacing: 0,
      lineSpacingType: "PERCENT",
      lineSpacingVal: 115,
      borderFillIDRef: "1",
      keepWithNext: "0",
      keepLines: "0",
      pageBreakBefore: "0",
      tabPrIDRef: "1",
      tabStops: []
    }), this.paraShapes.set("1", {
      id: "1",
      align: "LEFT",
      heading: "NONE",
      level: "0",
      headingIdRef: "0",
      leftMargin: 0,
      rightMargin: 0,
      indent: 0,
      prevSpacing: 0,
      nextSpacing: 0,
      lineSpacingType: "PERCENT",
      lineSpacingVal: 115,
      borderFillIDRef: "1",
      keepWithNext: "0",
      keepLines: "0",
      pageBreakBefore: "0",
      tabPrIDRef: "1",
      tabStops: []
    });
  }
}
function Mt(t) {
  return t ? String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;") : "";
}
function Ii(t) {
  return vi(t) ? "FCAT_FIXED" : t.includes("고딕") || t.includes("Gothic") || t.includes("Sans") || t.includes("Arial") || t.includes("Helvetica") || t.includes("Verdana") || t.includes("Tahoma") || t.includes("돋움") ? "FCAT_GOTHIC" : ["바탕", "명조", "Batang", "Times", "Serif", "Georgia"].some((e) => t.includes(e)) ? "FCAT_MYEONGJO" : "FCAT_UNKNOWN";
}
function vi(t) {
  return [
    "Courier New",
    "Consolas",
    "Lucida Console",
    "Monaco",
    "Menlo",
    "monospace",
    "Ubuntu Mono",
    "Source Code Pro",
    "Fira Code",
    "JetBrains Mono",
    "D2Coding",
    "NanumGothicCoding"
  ].some((e) => t.toLowerCase().includes(e.toLowerCase()));
}
function Ni(t, r) {
  let e = 0, n = 0;
  try {
    if (r === "png") {
      const i = new DataView(t.buffer, t.byteOffset, t.byteLength);
      i.getUint32(0) === 2303741511 && i.getUint32(4) === 218765834 && (e = i.getUint32(16), n = i.getUint32(20));
    } else if (r === "jpg" || r === "jpeg") {
      const i = new DataView(t.buffer, t.byteOffset, t.byteLength);
      let a = 2;
      for (; a < i.byteLength && !(a + 1 >= i.byteLength); ) {
        const s = i.getUint16(a);
        if (a += 2, s === 65472 || s === 65474) {
          n = i.getUint16(a + 3), e = i.getUint16(a + 5);
          break;
        } else if ((s & 65280) === 65280 && s !== 65535) {
          const o = i.getUint16(a);
          a += o;
        } else
          break;
      }
    }
  } catch (i) {
    console.warn("이미지 크기 추출 실패 (기본값 사용):", i);
  }
  return { width: e || 100, height: n || 100 };
}
class Ri {
  constructor(r) {
    ce(this, "state");
    this.state = r;
  }
  async parseMetadata(r) {
    var e;
    try {
      const n = await ((e = r.file("docProps/core.xml")) == null ? void 0 : e.async("text"));
      if (n) {
        const a = new DOMParser().parseFromString(n, "text/xml"), s = (o) => {
          var h;
          return ((h = a.getElementsByTagName(o)[0]) == null ? void 0 : h.textContent) || "";
        };
        this.state.metadata.title = s("dc:title") || s("title"), this.state.metadata.creator = s("dc:creator") || s("creator"), this.state.metadata.subject = s("dc:subject") || s("subject"), this.state.metadata.description = s("dc:description") || s("description"), this.state.metadata.keywords = s("cp:keywords") || s("keywords"), this.state.metadata.lastModifiedBy = s("cp:lastModifiedBy") || s("lastModifiedBy"), this.state.metadata.revision = s("cp:revision") || s("revision"), this.state.metadata.created = s("dcterms:created") || s("created"), this.state.metadata.modified = s("dcterms:modified") || s("modified");
      }
    } catch (n) {
      console.warn("메타데이터 파싱 실패:", n);
    }
  }
  async parseRelationships(r) {
    var e;
    try {
      const n = await ((e = r.file("word/_rels/document.xml.rels")) == null ? void 0 : e.async("text"));
      if (n) {
        const s = new DOMParser().parseFromString(n, "text/xml").getElementsByTagName("Relationship");
        for (let o = 0; o < s.length; o++) {
          const h = s[o], l = h.getAttribute("Id"), f = h.getAttribute("Target"), d = h.getAttribute("Type");
          l && f && d && this.state.relationships.set(l, { target: f, type: d });
        }
      }
    } catch (n) {
      console.warn("관계 파싱 실패:", n);
    }
  }
  async parseImages(r) {
    var e;
    for (const [n, i] of this.state.relationships)
      if ((e = i.type) != null && e.includes("image"))
        try {
          let a = i.target;
          a.startsWith("/") && (a = a.slice(1));
          const s = a.startsWith("word/") ? a : `word/${a}`, o = r.file(s);
          if (o) {
            const h = await o.async("uint8array"), l = (s.split(".").pop() || "").toLowerCase(), f = `BIN${String(this.state.imageCounter).padStart(4, "0")}`, d = Ni(h, l);
            this.state.images.set(n, {
              data: h,
              ext: l === "jpeg" ? "jpg" : l,
              // Standardize extension
              manifestId: f,
              path: `BinData/${f}.${l === "jpeg" ? "jpg" : l}`,
              width: d.width,
              height: d.height,
              binDataId: this.state.imageCounter
            }), this.state.imageCounter++;
          }
        } catch (a) {
          console.warn(`이미지 ${n} 파싱 실패:`, a);
        }
    console.log(`✅ ${this.state.images.size}개 이미지 파싱 완료`);
  }
  async parseStyles(r) {
    var e;
    try {
      const n = await ((e = r.file("word/styles.xml")) == null ? void 0 : e.async("text"));
      if (n) {
        const s = new DOMParser().parseFromString(n, "text/xml").getElementsByTagNameNS(this.state.ns.w, "style");
        for (let o = 0; o < s.length; o++) {
          const h = s[o], l = h.getAttribute("w:styleId");
          if (l) {
            let f = this.state.docStyles.get(l) || {};
            const d = h.getElementsByTagNameNS(this.state.ns.w, "pPr")[0];
            if (d) {
              const y = d.getElementsByTagNameNS(this.state.ns.w, "jc")[0];
              y && (f.align = y.getAttribute("w:val") || void 0);
              const m = d.getElementsByTagNameNS(this.state.ns.w, "ind")[0];
              m && (f.ind = {}, m.getAttribute("w:left") && (f.ind.left = m.getAttribute("w:left") || void 0), m.getAttribute("w:right") && (f.ind.right = m.getAttribute("w:right") || void 0), m.getAttribute("w:firstLine") && (f.ind.firstLine = m.getAttribute("w:firstLine") || void 0), m.getAttribute("w:hanging") && (f.ind.hanging = m.getAttribute("w:hanging") || void 0));
              const A = d.getElementsByTagNameNS(this.state.ns.w, "spacing")[0];
              A && (f.spacing = {}, A.getAttribute("w:before") && (f.spacing.before = A.getAttribute("w:before") || void 0), A.getAttribute("w:after") && (f.spacing.after = A.getAttribute("w:after") || void 0), A.getAttribute("w:line") && (f.spacing.line = A.getAttribute("w:line") || void 0), A.getAttribute("w:lineRule") && (f.spacing.lineRule = A.getAttribute("w:lineRule") || void 0));
            }
            const u = h.getElementsByTagNameNS(this.state.ns.w, "rPr")[0];
            if (u) {
              f.rPr || (f.rPr = {});
              const y = u.getElementsByTagNameNS(this.state.ns.w, "sz")[0];
              y && (f.rPr.sz = y.getAttribute("w:val") || void 0);
              const m = u.getElementsByTagNameNS(this.state.ns.w, "color")[0];
              m && (f.rPr.color = m.getAttribute("w:val") || void 0);
            }
            const c = h.getElementsByTagNameNS(this.state.ns.w, "tblPr")[0];
            if (c) {
              const y = c.getElementsByTagNameNS(this.state.ns.w, "tblBorders")[0];
              y && (f.tblBorders = y);
            }
            const g = h.getElementsByTagNameNS(this.state.ns.w, "tcPr")[0];
            if (g) {
              const y = g.getElementsByTagNameNS(this.state.ns.w, "tcBorders")[0];
              y && (f.tcBorders = y);
              const m = g.getElementsByTagNameNS(this.state.ns.w, "shd")[0];
              m && (f.shd = m);
            }
            const p = h.getElementsByTagNameNS(this.state.ns.w, "basedOn")[0];
            p && !f.basedOn && (f.basedOn = p.getAttribute("w:val") || void 0), this.state.docStyles.set(l, f);
          }
        }
      }
    } catch (n) {
      console.warn("스타일 파싱 실패:", n);
    }
  }
  async parseNumbering(r) {
    var e, n, i, a, s;
    try {
      const o = await ((e = r.file("word/numbering.xml")) == null ? void 0 : e.async("text"));
      if (o) {
        const l = new DOMParser().parseFromString(o, "text/xml"), f = l.getElementsByTagNameNS(this.state.ns.w, "abstractNum");
        for (let u = 0; u < f.length; u++) {
          const c = f[u], g = c.getAttribute("w:abstractNumId");
          if (g) {
            const p = /* @__PURE__ */ new Map(), y = c.getElementsByTagNameNS(this.state.ns.w, "lvl");
            for (let m = 0; m < y.length; m++) {
              const A = y[m], E = A.getAttribute("w:ilvl"), x = parseInt(((n = A.getElementsByTagNameNS(this.state.ns.w, "start")[0]) == null ? void 0 : n.getAttribute("w:val")) || "1"), S = ((i = A.getElementsByTagNameNS(this.state.ns.w, "numFmt")[0]) == null ? void 0 : i.getAttribute("w:val")) || "decimal", R = ((a = A.getElementsByTagNameNS(this.state.ns.w, "lvlText")[0]) == null ? void 0 : a.getAttribute("w:val")) || "";
              let v = null, O = null;
              const C = A.getElementsByTagNameNS(this.state.ns.w, "pPr")[0];
              if (C) {
                const B = C.getElementsByTagNameNS(this.state.ns.w, "ind")[0];
                B && (v = B.getAttribute("w:left") || B.getAttribute("w:start"), O = B.getAttribute("w:hanging"));
              }
              E && p.set(E, { start: x, numFmt: S, lvlText: R, lvlLeft: v, lvlHanging: O });
            }
            this.state.abstractNumMap.set(g, p);
          }
        }
        const d = l.getElementsByTagNameNS(this.state.ns.w, "num");
        for (let u = 0; u < d.length; u++) {
          const c = d[u], g = c.getAttribute("w:numId"), p = (s = c.getElementsByTagNameNS(this.state.ns.w, "abstractNumId")[0]) == null ? void 0 : s.getAttribute("w:val");
          g && p && this.state.numberingMap.set(g, p);
        }
      }
    } catch (o) {
      console.warn("번호 매기기 파싱 실패:", o);
    }
  }
  async parseFootnotes(r) {
    var e;
    try {
      const n = await ((e = r.file("word/footnotes.xml")) == null ? void 0 : e.async("text"));
      if (n) {
        const s = new DOMParser().parseFromString(n, "text/xml").getElementsByTagNameNS(this.state.ns.w, "footnote");
        for (let o = 0; o < s.length; o++) {
          const h = s[o], l = h.getAttribute("w:id");
          l && this.state.footnoteMap.set(l, h);
        }
      }
    } catch (n) {
      console.warn("각주 파싱 실패:", n);
    }
  }
  async parseEndnotes(r) {
    var e;
    try {
      const n = await ((e = r.file("word/endnotes.xml")) == null ? void 0 : e.async("text"));
      if (n) {
        const s = new DOMParser().parseFromString(n, "text/xml").getElementsByTagNameNS(this.state.ns.w, "endnote");
        for (let o = 0; o < s.length; o++) {
          const h = s[o], l = h.getAttribute("w:id");
          l && this.state.endnoteMap.set(l, h);
        }
      }
    } catch (n) {
      console.warn("미주 파싱 실패:", n);
    }
  }
  parsePageSetup(r) {
    const e = r.getElementsByTagNameNS(this.state.ns.w, "sectPr")[0];
    if (!e) return;
    const n = e.getElementsByTagNameNS(this.state.ns.w, "pgSz")[0];
    if (n) {
      const s = parseInt(n.getAttribute("w:w") || "11906"), o = parseInt(n.getAttribute("w:h") || "16838");
      this.state.PAGE_WIDTH = s * 5, this.state.PAGE_HEIGHT = o * 5, this.state.TEXT_WIDTH = this.state.PAGE_WIDTH - this.state.MARGIN_LEFT - this.state.MARGIN_RIGHT;
    }
    const i = e.getElementsByTagNameNS(this.state.ns.w, "pgMar")[0];
    if (i) {
      const s = parseInt(i.getAttribute("w:top") || "1440"), o = parseInt(i.getAttribute("w:bottom") || "1440"), h = parseInt(i.getAttribute("w:left") || "1800"), l = parseInt(i.getAttribute("w:right") || "1800"), f = parseInt(i.getAttribute("w:header") || "720"), d = parseInt(i.getAttribute("w:footer") || "720"), u = parseInt(i.getAttribute("w:gutter") || "0");
      this.state.MARGIN_TOP = s * 5, this.state.MARGIN_BOTTOM = o * 5, this.state.MARGIN_LEFT = h * 5, this.state.MARGIN_RIGHT = l * 5, this.state.HEADER_MARGIN = f * 5, this.state.FOOTER_MARGIN = d * 5, this.state.GUTTER_MARGIN = u * 5, this.state.TEXT_WIDTH = this.state.PAGE_WIDTH - this.state.MARGIN_LEFT - this.state.MARGIN_RIGHT - this.state.GUTTER_MARGIN;
    }
    const a = e.getElementsByTagNameNS(this.state.ns.w, "cols")[0];
    if (a) {
      this.state.COL_COUNT = parseInt(a.getAttribute("w:num") || "1");
      const s = parseInt(a.getAttribute("w:space") || "720");
      this.state.COL_GAP = s * 5, this.state.COL_TYPE = this.state.COL_COUNT > 1 ? "NEWSPAPER" : "ONE";
    }
  }
}
class Ci {
  constructor(r) {
    ce(this, "state");
    this.state = r;
  }
  createFontFaces() {
    const r = ["HANGUL", "LATIN", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER"];
    let e = `<hh:fontfaces itemCnt="${r.length}">`;
    for (const n of r) {
      const i = this.state.langFontFaces[n];
      e += `<hh:fontface lang="${n}" fontCnt="${i.size}">`;
      for (const [a, s] of i) {
        const o = Ii(a);
        e += `<hh:font id="${s}" face="${a}" type="TTF" isEmbedded="0">`, e += `<hh:typeInfo familyType="${o}" weight="0" proportion="0" contrast="0" strokeVariation="0" armStyle="0" letterform="0" midline="0" xHeight="0"/>`, e += "</hh:font>";
      }
      e += "</hh:fontface>";
    }
    return e += "</hh:fontfaces>", e;
  }
  createSection(r) {
    let e = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec xmlns:ha="${this.state.hwpxNs.ha}" xmlns:hp="${this.state.hwpxNs.hp}" xmlns:hp10="${this.state.hwpxNs.hp10}" xmlns:hs="${this.state.hwpxNs.hs}" xmlns:hc="${this.state.hwpxNs.hc}" xmlns:hh="${this.state.hwpxNs.hh}" xmlns:hhs="${this.state.hwpxNs.hhs}" xmlns:hm="${this.state.hwpxNs.hm}" xmlns:hpf="${this.state.hwpxNs.hpf}" xmlns:dc="${this.state.hwpxNs.dc}" xmlns:opf="${this.state.hwpxNs.opf}" xmlns:ooxmlchart="${this.state.hwpxNs.ooxmlchart}" xmlns:epub="${this.state.hwpxNs.epub}" xmlns:config="${this.state.hwpxNs.config}">`;
    const n = r.getElementsByTagNameNS(this.state.ns.w, "body")[0];
    if (n) {
      const i = Array.from(n.children).filter(
        (s) => s.localName === "p" || s.localName === "tbl" || s.localName === "sdt"
      ), a = i.length;
      i.forEach((s, o) => {
        const h = o === a - 1;
        if (s.localName === "p")
          e += this.convertParagraph(s, h);
        else if (s.localName === "tbl")
          e += this.convertTable(s);
        else if (s.localName === "sdt") {
          const l = s.getElementsByTagNameNS(this.state.ns.w, "sdtContent")[0];
          if (l) {
            const f = Array.from(l.children).filter(
              (u) => u.localName === "p" || u.localName === "tbl"
            ), d = f.length;
            f.forEach((u, c) => {
              const g = h && c === d - 1;
              u.localName === "p" ? e += this.convertParagraph(u, g) : u.localName === "tbl" && (e += this.convertTable(u));
            });
          }
        }
      });
    } else {
      const i = this.state.idCounter++;
      e += `<hp:p id="${i}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="2">${this.createSecPr()}<hp:t></hp:t></hp:run><hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1265" textheight="1265" baseline="1032" spacing="188" horzpos="0" horzsize="${this.state.TEXT_WIDTH}" flags="393216"/></hp:linesegarray></hp:p>`;
    }
    return e += "</hs:sec>", e;
  }
  createSecPr() {
    return `<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="${this.state.COL_GAP}" tabStop="8000" outlineShapeIDRef="1" memoShapeIDRef="1" textVerticalWidthHead="0" masterPageCnt="0"><hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/><hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/><hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/><hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/><hp:pagePr landscape="WIDELY" width="${this.state.PAGE_WIDTH}" height="${this.state.PAGE_HEIGHT}" gutterType="LEFT_ONLY"><hp:margin header="${this.state.HEADER_MARGIN}" footer="${this.state.FOOTER_MARGIN}" gutter="${this.state.GUTTER_MARGIN}" left="${this.state.MARGIN_LEFT}" right="${this.state.MARGIN_RIGHT}" top="${this.state.MARGIN_TOP}" bottom="${this.state.MARGIN_BOTTOM}"/></hp:pagePr><hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supscript="1"/><hp:noteLine length="-1" type="SOLID" width="0.25 mm" color="#000000"/><hp:noteSpacing betweenNotes="283" belowLine="0" aboveLine="1000"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="EACH_COLUMN" beneathText="0"/></hp:footNotePr><hp:endNotePr><hp:autoNumFormat type="ROMAN_SMALL" userChar="" prefixChar="" suffixChar="" supscript="1"/><hp:noteLine length="-1" type="SOLID" width="0.25 mm" color="#000000"/><hp:noteSpacing betweenNotes="0" belowLine="0" aboveLine="1000"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="END_OF_DOCUMENT" beneathText="0"/></hp:endNotePr><hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill><hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill><hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill></hp:secPr><hp:ctrl><hp:colPr id="" type="${this.state.COL_TYPE}" layout="LEFT" colCount="${this.state.COL_COUNT}" sameSz="1" sameGap="0"/></hp:ctrl>`;
  }
  convertParagraph(r, e = !1) {
    var R, v;
    const n = this.state.idCounter++, i = this.getParaShapeId(r, e);
    let a = null;
    const s = r.getElementsByTagNameNS(this.state.ns.w, "pPr")[0];
    if (s) {
      const O = s.getElementsByTagNameNS(this.state.ns.w, "pStyle")[0];
      O && (a = O.getAttribute("w:val"));
    }
    const o = this.getParagraphFontSize(r), h = Math.round(o * this.state.HWPUNIT_PER_INCH / 72), l = this.state.paraShapes.get(i), f = l && l.lineSpacingType === "PERCENT" ? l.lineSpacingVal : 115, d = h, u = Math.round(d * 0.85), c = Math.round(d * (f - 100) / 100), g = d + c;
    let p = "0";
    if (s) {
      const O = s.getElementsByTagNameNS(this.state.ns.w, "pageBreakBefore")[0];
      O && O.getAttribute("w:val") !== "false" && (p = "1");
      const C = s.getElementsByTagNameNS(this.state.ns.w, "rPr")[0];
      this.getCharShapeId(C || null, o, !1, a);
    } else
      this.getCharShapeId(null, o, !1, a);
    let y = "0";
    a && this.state.docxStyleToHwpxId && this.state.docxStyleToHwpxId[a] !== void 0 && (y = String(this.state.docxStyleToHwpxId[a]));
    let m = `<hp:p id="${n}" paraPrIDRef="${i}" styleIDRef="${y}" pageBreak="${p}" columnBreak="0" merged="0">`;
    this.state.isFirstParagraph && (m += `<hp:run charPrIDRef="0">${this.createSecPr()}</hp:run>`, this.state.isFirstParagraph = !1);
    const A = s == null ? void 0 : s.getElementsByTagNameNS(this.state.ns.w, "numPr")[0];
    if (A) {
      const O = (R = A.getElementsByTagNameNS(this.state.ns.w, "numId")[0]) == null ? void 0 : R.getAttribute("w:val"), C = (v = A.getElementsByTagNameNS(this.state.ns.w, "ilvl")[0]) == null ? void 0 : v.getAttribute("w:val");
      O && C && this.getListPrefix(O, C);
    }
    const E = r.childNodes;
    for (let O = 0; O < E.length; O++) {
      const C = E[O];
      if (C.nodeName === "w:r")
        m += this.convertRun(C, o, a);
      else if (C.nodeName === "w:hyperlink")
        for (let B = 0; B < C.childNodes.length; B++) {
          const Z = C.childNodes[B];
          Z.nodeName === "w:r" && (m += this.convertRun(Z, o, a));
        }
    }
    const x = l && l.prevSpacing ? l.prevSpacing : 0, S = l && l.nextSpacing ? l.nextSpacing : 0;
    return this.state.currentVertPos += x, m += `<hp:linesegarray><hp:lineseg textpos="0" vertpos="${this.state.currentVertPos}" vertsize="${g}" textheight="${d}" baseline="${u}" spacing="${c}" horzpos="0" horzsize="${this.state.TEXT_WIDTH}" flags="393216"/></hp:linesegarray></hp:p>`, this.state.currentVertPos += g + S, m;
  }
  convertRun(r, e, n = null) {
    const i = r.getElementsByTagNameNS(this.state.ns.w, "rPr")[0];
    let a = !1;
    r.parentNode && r.parentNode.nodeName === "w:hyperlink" && (a = !0);
    const s = this.getCharShapeId(i, e, a, n);
    let o = "", h = "", l = !1;
    const f = () => {
      (h.length > 0 || l) && (o += `<hp:run charPrIDRef="${s}">`, o += h, o += "</hp:run>", h = "", l = !1);
    };
    for (let d = 0; d < r.childNodes.length; d++) {
      const u = r.childNodes[d];
      if (u.nodeName === "w:t")
        h += `<hp:t>${Mt(u.textContent || "")}</hp:t>`, l = !0;
      else if (u.nodeName === "w:tab")
        h += "<hp:t>	</hp:t>", l = !0;
      else if (u.nodeName === "w:br")
        u.getAttribute("w:type") !== "page" && (h += `<hp:t>
</hp:t>`, l = !0);
      else if (u.nodeName === "w:footnoteReference") {
        const c = u.getAttribute("w:id");
        c && (h += this.convertFootnote(c, "FOOTNOTE"), l = !0);
      } else if (u.nodeName === "w:endnoteReference") {
        const c = u.getAttribute("w:id");
        c && (h += this.convertFootnote(c, "ENDNOTE"), l = !0);
      } else if (u.nodeName === "w:drawing")
        f(), o += this.convertDrawing(u);
      else if (u.localName === "AlternateContent" || u.nodeName.includes("AlternateContent")) {
        const c = u.getElementsByTagName("*"), g = Array.from(c).filter((p) => p.localName === "drawing");
        g.length > 0 && (f(), o += this.convertDrawing(g[0]));
      }
    }
    return f(), o;
  }
  convertFootnote(r, e) {
    const i = (e === "FOOTNOTE" ? this.state.footnoteMap : this.state.endnoteMap).get(r);
    if (!i) return "";
    const a = e === "FOOTNOTE", s = a ? ++this.state.footnoteNumber : ++this.state.endnoteNumber, o = Math.floor(Math.random() * 1e7), h = i.getElementsByTagNameNS(this.state.ns.w, "p");
    let l = "", f = !0;
    for (let c = 0; c < h.length; c++) {
      const g = h[c];
      if (f) {
        const p = this.convertParagraph(g), m = `<hp:ctrl><hp:autoNum num="${s}" numType="${a ? "FOOTNOTE" : "ENDNOTE"}"><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supscript="1"/></hp:autoNum></hp:ctrl>`;
        l += p.replace(/(<hp:run[^>]*>)/, `$1${m}`), f = !1;
      } else
        l += this.convertParagraph(g);
    }
    const d = a ? "hp:footNote" : "hp:endNote";
    return `<hp:ctrl><${d} number="${s}" instId="${o}"${a ? "" : ' flag="3"'}><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">${l}</hp:subList></${d}></hp:ctrl>`;
  }
  getParaShapeId(r, e = !1) {
    var O, C;
    if (!r || typeof r.getElementsByTagNameNS != "function") return "1";
    const n = r.getElementsByTagNameNS(this.state.ns.w, "pPr")[0];
    let i = "LEFT", a = "NONE", s = "0", o = "0", h = 0, l = 0, f = 0, d = 0, u = 0, c = "PERCENT", g = 115, p = "0", y = "0", m = "0", A = "1", E = !1, x = !1, S = null;
    if (n) {
      let B = null;
      const Z = n.getElementsByTagNameNS(this.state.ns.w, "jc")[0];
      if (Z)
        B = Z.getAttribute("w:val");
      else {
        const K = n.getElementsByTagNameNS(this.state.ns.w, "pStyle")[0];
        if (K) {
          S = K.getAttribute("w:val");
          let se = S;
          for (; se; ) {
            const P = this.state.docStyles.get(se);
            if (!P) break;
            if (P.align) {
              B = P.align;
              break;
            }
            se = P.basedOn;
          }
        }
      }
      B === "center" ? i = "CENTER" : B === "right" ? i = "RIGHT" : B === "both" ? i = "JUSTIFY" : B === "distribute" && (i = "DISTRIBUTE");
      const T = n.getElementsByTagNameNS(this.state.ns.w, "ind")[0];
      if (T) {
        E = !0;
        const K = T.getAttribute("w:left") || T.getAttribute("w:start"), se = T.getAttribute("w:right") || T.getAttribute("w:end"), P = T.getAttribute("w:firstLine"), G = T.getAttribute("w:hanging");
        K && (h = parseInt(K) * 5), se && (l = parseInt(se) * 5), P ? f = parseInt(P) * 5 : G && (f = -parseInt(G) * 5);
      }
      const j = n.getElementsByTagNameNS(this.state.ns.w, "numPr")[0];
      if (j) {
        const K = (O = j.getElementsByTagNameNS(this.state.ns.w, "numId")[0]) == null ? void 0 : O.getAttribute("w:val"), se = ((C = j.getElementsByTagNameNS(this.state.ns.w, "ilvl")[0]) == null ? void 0 : C.getAttribute("w:val")) || "0";
        if (K && K !== "0") {
          const P = this.state.numberingMap.get(K);
          if (P) {
            let G = 1;
            for (const [U] of this.state.abstractNumMap) {
              if (U === P) break;
              G++;
            }
            if (a = "NUMBER", o = String(G), s = String(parseInt(se) + 1), !E) {
              const U = this.state.abstractNumMap.get(P);
              if (U) {
                const Q = U.get(se);
                Q && Q.lvlLeft ? (h = parseInt(Q.lvlLeft) * 5, Q.lvlHanging && (f = -parseInt(Q.lvlHanging) * 5), E = !0) : (h = (parseInt(se) + 1) * 800 * 5, f = -400 * 5, E = !0);
              }
            }
          }
        } else K === "0" && (a = "NONE");
      }
      const _ = n.getElementsByTagNameNS(this.state.ns.w, "spacing")[0];
      if (_) {
        x = !0;
        const K = _.getAttribute("w:before"), se = _.getAttribute("w:after");
        K && (d = parseInt(K) * 5), se && (u = 0);
        const P = _.getAttribute("w:line"), G = _.getAttribute("w:lineRule");
        P && (G === "auto" || !G ? g = Math.round(parseInt(P) / 240 * 100) : G === "exact" ? (c = "FIXED", g = parseInt(P) * 5) : G === "atLeast" && (c = "ATLEAST", g = parseInt(P) * 5));
      }
    }
    if ((!E || !x) && S) {
      let B = S;
      for (; B; ) {
        const Z = this.state.docStyles.get(B);
        if (!Z) break;
        if (!E && Z.ind && (Z.ind.left && (h = parseInt(Z.ind.left) * 5), Z.ind.right && (l = parseInt(Z.ind.right) * 5), Z.ind.firstLine && (f = parseInt(Z.ind.firstLine) * 5), Z.ind.hanging && (f = -parseInt(Z.ind.hanging) * 5), E = !0), !x && Z.spacing) {
          if (Z.spacing.before && (d = parseInt(Z.spacing.before) * 5), Z.spacing.after && (u = parseInt(Z.spacing.after) * 5), Z.spacing.line) {
            const T = Z.spacing.lineRule;
            T === "auto" || !T ? g = Math.round(parseInt(Z.spacing.line) / 240 * 100) : T === "exact" ? (c = "FIXED", g = parseInt(Z.spacing.line) * 5) : T === "atLeast" && (c = "ATLEAST", g = parseInt(Z.spacing.line) * 5);
          }
          x = !0;
        }
        if (E && x) break;
        B = Z.basedOn;
      }
    }
    if (n) {
      const B = n.getElementsByTagNameNS(this.state.ns.w, "outlineLvl")[0];
      B && (s = B.getAttribute("w:val") || "0");
      const Z = n.getElementsByTagNameNS(this.state.ns.w, "keepNext")[0];
      Z && Z.getAttribute("w:val") !== "false" && Z.getAttribute("w:val") !== "0" && (p = "1");
      const T = n.getElementsByTagNameNS(this.state.ns.w, "keepLines")[0];
      T && T.getAttribute("w:val") !== "false" && T.getAttribute("w:val") !== "0" && (y = "1");
      const j = n.getElementsByTagNameNS(this.state.ns.w, "pageBreakBefore")[0];
      j && j.getAttribute("w:val") !== "false" && j.getAttribute("w:val") !== "0" && (m = "1");
      const _ = n.getElementsByTagNameNS(this.state.ns.w, "pBdr")[0], K = n.getElementsByTagNameNS(this.state.ns.w, "shd")[0];
      if (_ || K) {
        const se = (le) => {
          if (!le) return { type: "NONE", width: "0.1 mm", color: "#000000" };
          const ne = le.getAttribute("w:val");
          if (ne === "none" || ne === "nil") return { type: "NONE", width: "0.1 mm", color: "#000000" };
          const ye = le.getAttribute("w:sz");
          let Te = "0.1 mm";
          ye && (Te = (parseInt(ye) / 8 * 0.3528).toFixed(2) + " mm");
          const _e = le.getAttribute("w:color");
          let Se = "#000000";
          _e && _e !== "auto" && (Se = `#${_e}`);
          let ve = "SOLID";
          return ne === "double" ? ve = "DOUBLE" : ne === "dotted" ? ve = "DOT" : ne === "dashed" && (ve = "DASH"), { type: ve, width: Te, color: Se };
        }, P = se(_ == null ? void 0 : _.getElementsByTagNameNS(this.state.ns.w, "left")[0]), G = se(_ == null ? void 0 : _.getElementsByTagNameNS(this.state.ns.w, "right")[0]), U = se(_ == null ? void 0 : _.getElementsByTagNameNS(this.state.ns.w, "top")[0]), Q = se(_ == null ? void 0 : _.getElementsByTagNameNS(this.state.ns.w, "bottom")[0]);
        let M = "#FFFFFF";
        if (K) {
          const le = K.getAttribute("w:fill");
          le && le !== "auto" && (M = `#${le}`);
        }
        const X = `${P.type}_${P.width}_${P.color}|${G.type}_${G.width}_${G.color}|${U.type}_${U.width}_${U.color}|${Q.type}_${Q.width}_${Q.color}|${M}`;
        let de = null;
        for (const [le, ne] of this.state.borderFills) {
          const ye = `${ne.leftBorder.type}_${ne.leftBorder.width}_${ne.leftBorder.color}|${ne.rightBorder.type}_${ne.rightBorder.width}_${ne.rightBorder.color}|${ne.topBorder.type}_${ne.topBorder.width}_${ne.topBorder.color}|${ne.bottomBorder.type}_${ne.bottomBorder.width}_${ne.bottomBorder.color}|${ne.backColor || "#FFFFFF"}`;
          if (X === ye) {
            de = le;
            break;
          }
        }
        if (de)
          A = de;
        else {
          const le = String(this.state.borderFillIdCounter++);
          this.state.borderFills.set(le, {
            id: le,
            leftBorder: P,
            rightBorder: G,
            topBorder: U,
            bottomBorder: Q,
            diagonal: null,
            slash: { type: "NONE", width: "0.1 mm", color: "#000000" },
            backSlash: { type: "NONE", width: "0.1 mm", color: "#000000" },
            backColor: M
          }), A = le;
        }
      }
    }
    e && (u = 0);
    const R = `${i}_${h}_${l}_${f}_${d}_${u}_${c}_${g}_${a}_${s}_${o}_${A}_${p}_${y}_${m}`;
    for (const [B, Z] of this.state.paraShapes)
      if (Z._key === R) return B;
    const v = String(this.state.paraShapes.size);
    return this.state.paraShapes.set(v, {
      id: v,
      align: i,
      heading: a,
      level: s,
      headingIdRef: o,
      leftMargin: h,
      rightMargin: l,
      indent: f,
      prevSpacing: d,
      nextSpacing: u,
      lineSpacingType: c,
      lineSpacingVal: g,
      borderFillIDRef: A,
      keepWithNext: p,
      keepLines: y,
      pageBreakBefore: m,
      tabPrIDRef: "1",
      tabStops: [],
      _key: R
    }), v;
  }
  getCharShapeId(r, e, n = !1, i = null) {
    let a = !1, s = !1, o = n, h = !1, l = !1, f = !1, d = n ? "#0000FF" : "#000000", u = "none", c = "1", g = e ? Math.round(e * 100) : 1e3, p = null, y = !1;
    if (r) {
      r.getElementsByTagNameNS(this.state.ns.w, "b").length > 0 && r.getElementsByTagNameNS(this.state.ns.w, "b")[0].getAttribute("w:val") !== "0" && (a = !0), r.getElementsByTagNameNS(this.state.ns.w, "i").length > 0 && r.getElementsByTagNameNS(this.state.ns.w, "i")[0].getAttribute("w:val") !== "0" && (s = !0), r.getElementsByTagNameNS(this.state.ns.w, "u").length > 0 && r.getElementsByTagNameNS(this.state.ns.w, "u")[0].getAttribute("w:val") !== "none" && (o = !0), r.getElementsByTagNameNS(this.state.ns.w, "strike").length > 0 && r.getElementsByTagNameNS(this.state.ns.w, "strike")[0].getAttribute("w:val") !== "0" && (h = !0), r.getElementsByTagNameNS(this.state.ns.w, "dstrike").length > 0 && r.getElementsByTagNameNS(this.state.ns.w, "dstrike")[0].getAttribute("w:val") !== "0" && (h = !0);
      const S = r.getElementsByTagNameNS(this.state.ns.w, "vertAlign")[0];
      if (S) {
        const T = S.getAttribute("w:val");
        T === "superscript" ? l = !0 : T === "subscript" && (f = !0);
      }
      const R = r.getElementsByTagNameNS(this.state.ns.w, "color")[0];
      if (R) {
        const T = R.getAttribute("w:val");
        T && T !== "auto" && (d = `#${T}`);
      }
      const v = r.getElementsByTagNameNS(this.state.ns.w, "sz")[0];
      if (v) {
        const T = parseInt(v.getAttribute("w:val") || "0");
        !isNaN(T) && T > 0 && (g = T * 50, y = !0);
      }
      const O = r.getElementsByTagNameNS(this.state.ns.w, "highlight")[0];
      if (O) {
        const T = O.getAttribute("w:val") || "", j = {
          black: "#000000",
          blue: "#0000FF",
          cyan: "#00FFFF",
          green: "#008000",
          magenta: "#FF00FF",
          red: "#FF0000",
          yellow: "#FFFF00",
          white: "#FFFFFF",
          darkBlue: "#00008B",
          darkCyan: "#008B8B",
          darkGreen: "#006400",
          darkMagenta: "#8B008B",
          darkRed: "#8B0000",
          darkYellow: "#808000",
          darkGray: "#A9A9A9",
          lightGray: "#D3D3D3",
          none: "none"
        };
        j[T] && j[T] !== "none" && (u = j[T], c = this.getOrCreateHighlightBorderFill(u));
      }
      const C = r.getElementsByTagNameNS(this.state.ns.w, "shd")[0];
      if (C) {
        const T = C.getAttribute("w:fill");
        T && T !== "auto" && (T === "000000" || T === "black" ? (c = "6", d === "#000000" && (d = "#FFFFFF")) : (u = `#${T}`, c = this.getOrCreateHighlightBorderFill(u)));
      }
      r.getElementsByTagNameNS(this.state.ns.w, "bdr")[0] && (c = "5");
      const Z = r.getElementsByTagNameNS(this.state.ns.w, "rFonts")[0];
      Z && (p = Z);
    }
    if (!y && i) {
      let S = i;
      for (; S; ) {
        const R = this.state.docStyles.get(S);
        if (!R) break;
        if (R.rPr && R.rPr.sz) {
          const v = parseInt(R.rPr.sz);
          if (!isNaN(v)) {
            g = v * 50;
            break;
          }
        }
        S = R.basedOn;
      }
    }
    if (d === "#000000" && !n && i) {
      let S = i;
      for (; S; ) {
        const R = this.state.docStyles.get(S);
        if (!R) break;
        if (R.rPr && R.rPr.color && R.rPr.color !== "auto") {
          d = `#${R.rPr.color}`;
          break;
        }
        S = R.basedOn;
      }
    }
    const m = this.registerFontsFromRFonts(p), A = g, E = `${a}_${s}_${o}_${h}_${l}_${f}_${d}_${u}_${c}_${A}_${m.hangulId}_${m.latinId}`;
    for (const [S, R] of this.state.charShapes)
      if (R._key === E) return S;
    const x = String(this.state.charShapes.size);
    return this.state.charShapes.set(x, {
      id: x,
      height: String(A),
      textColor: d,
      shadeColor: u,
      borderFillIDRef: c,
      bold: a,
      italic: s,
      underline: o,
      strike: h,
      supscript: l,
      subscript: f,
      fontIds: m,
      fontId: m.hangulId,
      _key: E
    }), x;
  }
  getOrCreateHighlightBorderFill(r) {
    for (const [n, i] of this.state.borderFills)
      if (i.backColor === r && i.leftBorder.type === "NONE" && i.rightBorder.type === "NONE" && i.topBorder.type === "NONE" && i.bottomBorder.type === "NONE")
        return n;
    const e = String(this.state.borderFillIdCounter++);
    return this.state.borderFills.set(e, {
      id: e,
      leftBorder: { type: "NONE", width: "0.1 mm", color: "#000000" },
      rightBorder: { type: "NONE", width: "0.1 mm", color: "#000000" },
      topBorder: { type: "NONE", width: "0.1 mm", color: "#000000" },
      bottomBorder: { type: "NONE", width: "0.1 mm", color: "#000000" },
      diagonal: null,
      slash: { type: "NONE", width: "0.1 mm", color: "#000000" },
      backSlash: { type: "NONE", width: "0.1 mm", color: "#000000" },
      backColor: r
    }), e;
  }
  registerFontsFromRFonts(r) {
    let e = "함초롬바탕", n = "함초롬바탕";
    if (r) {
      const i = r.getAttribute("w:eastAsia"), a = r.getAttribute("w:ascii"), s = r.getAttribute("w:hAnsi");
      i && (e = i), a ? n = a : s && (n = s);
    }
    return {
      hangulId: this.state.registerFontForLang("HANGUL", e),
      latinId: this.state.registerFontForLang("LATIN", n),
      hanjaId: this.state.registerFontForLang("HANJA", e),
      japaneseId: this.state.registerFontForLang("JAPANESE", e),
      otherId: this.state.registerFontForLang("OTHER", n),
      symbolId: this.state.registerFontForLang("SYMBOL", n),
      userId: this.state.registerFontForLang("USER", e)
    };
  }
  getParagraphFontSize(r) {
    const e = r.getElementsByTagNameNS(this.state.ns.w, "r");
    for (let i = 0; i < e.length; i++) {
      const a = e[i].getElementsByTagNameNS(this.state.ns.w, "rPr")[0];
      if (a) {
        const s = a.getElementsByTagNameNS(this.state.ns.w, "sz")[0];
        if (s) {
          const o = s.getAttribute("w:val");
          if (o) return parseInt(o) / 2;
        }
      }
    }
    const n = r.getElementsByTagNameNS(this.state.ns.w, "pPr")[0];
    if (n) {
      const i = n.getElementsByTagNameNS(this.state.ns.w, "rPr")[0];
      if (i) {
        const s = i.getElementsByTagNameNS(this.state.ns.w, "sz")[0];
        if (s) {
          const o = s.getAttribute("w:val");
          if (o) return parseInt(o) / 2;
        }
      }
      const a = n.getElementsByTagNameNS(this.state.ns.w, "pStyle")[0];
      if (a) {
        let s = a.getAttribute("w:val");
        for (; s; ) {
          const o = this.state.docStyles.get(s);
          if (!o) break;
          if (o.rPr && o.rPr.sz) {
            const h = parseInt(o.rPr.sz);
            if (!isNaN(h)) return h / 2;
          }
          s = o.basedOn;
        }
      }
    }
    return 10;
  }
  getListPrefix(r, e) {
    const n = this.state.numberingMap.get(r);
    if (!n) return "";
    const i = this.state.abstractNumMap.get(n);
    if (!i) return "";
    const a = i.get(e);
    if (!a) return "";
    const s = `${r}:${e}`;
    return this.state.listCounters && (this.state.listCounters.has(s) ? this.state.listCounters.set(s, this.state.listCounters.get(s) + 1) : this.state.listCounters.set(s, a.start || 1)), "";
  }
  getCellBorderFill(r, e, n, i, a, s, o = 1, h = 1, l = null, f = null) {
    const d = {
      left: { type: "SOLID", width: "0.1 mm", color: "#000000" },
      right: { type: "SOLID", width: "0.1 mm", color: "#000000" },
      top: { type: "SOLID", width: "0.1 mm", color: "#000000" },
      bottom: { type: "SOLID", width: "0.1 mm", color: "#000000" },
      slash: { type: "NONE", width: "0.1 mm", color: "#000000" },
      backSlash: { type: "NONE", width: "0.1 mm", color: "#000000" }
    }, u = (S) => {
      if (!S) return null;
      const R = S.getAttribute("w:val");
      if (R === "nil" || R === "none") return { type: "NONE", width: "0.1 mm", color: "#000000" };
      let v = "SOLID";
      R === "double" ? v = "DOUBLE" : R === "dashed" ? v = "DASH" : R === "dotted" && (v = "DOT");
      const C = parseInt(S.getAttribute("w:sz") || "4") / 8 * 0.3528, B = [0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1, 1.5, 2, 3, 4, 5];
      let Z = B[0], T = Math.abs(C - B[0]);
      for (let K = 1; K < B.length; K++) {
        const se = Math.abs(C - B[K]);
        se < T && (T = se, Z = B[K]);
      }
      C > 5 && (Z = 5);
      const j = S.getAttribute("w:color");
      let _ = "#000000";
      return j && j !== "auto" && (_ = `#${j}`), { type: v, width: `${Z} mm`, color: _ };
    }, c = (S, R, v) => {
      if (!S) return;
      const O = S.getElementsByTagNameNS(this.state.ns.w, R)[0];
      if (O) {
        const C = u(O);
        C && (d[v] = C);
      }
    };
    if (e) {
      const S = e.getElementsByTagNameNS(this.state.ns.w, "insideH")[0], R = e.getElementsByTagNameNS(this.state.ns.w, "insideV")[0], v = u(S), O = u(R);
      v && (d.top = { ...v }, d.bottom = { ...v }), O && (d.left = { ...O }, d.right = { ...O }), n === 0 && c(e, "top", "top"), n + o >= a && c(e, "bottom", "bottom"), i === 0 && c(e, "left", "left"), i + h >= s && c(e, "right", "right");
    }
    let g = r ? r.getElementsByTagNameNS(this.state.ns.w, "tcBorders")[0] : null;
    g || (g = l), g && (c(g, "top", "top"), c(g, "bottom", "bottom"), c(g, "left", "left"), c(g, "right", "right"), c(g, "tl2br", "slash"), c(g, "tr2bl", "backSlash"));
    let p = "#FFFFFF", y = r ? r.getElementsByTagNameNS(this.state.ns.w, "shd")[0] : null;
    if (y || (y = f), y) {
      const S = y.getAttribute("w:fill");
      S && S !== "auto" && (p = `#${S}`);
    }
    const m = {
      L: d.left,
      R: d.right,
      T: d.top,
      B: d.bottom,
      Sl: d.slash,
      Bs: d.backSlash,
      BG: p
    }, A = JSON.stringify(m);
    for (const [, S] of this.state.borderFills)
      if (S._key === A) return S;
    const E = String(this.state.borderFillIdCounter++), x = {
      id: E,
      leftBorder: d.left,
      rightBorder: d.right,
      topBorder: d.top,
      bottomBorder: d.bottom,
      diagonal: null,
      slash: d.slash,
      backSlash: d.backSlash,
      backColor: p,
      _key: A
    };
    return this.state.borderFills.set(E, x), x;
  }
  convertTable(r) {
    const e = this.state.tableIdCounter++, n = r.getElementsByTagNameNS(this.state.ns.w, "tblPr")[0];
    let i = null, a, s = { left: 283, right: 283, top: 283, bottom: 283 };
    if (n) {
      i = n.getElementsByTagNameNS(this.state.ns.w, "tblBorders")[0] || null;
      const S = n.getElementsByTagNameNS(this.state.ns.w, "tblStyle")[0];
      S && (a = S.getAttribute("w:val") || void 0);
      const R = n.getElementsByTagNameNS(this.state.ns.w, "tblCellMar")[0];
      if (R) {
        const v = (O, C) => {
          const B = R.getElementsByTagNameNS(this.state.ns.w, O)[0];
          return B ? parseInt(B.getAttribute("w:w") || "0") * 5 : C;
        };
        s.left = v("left", s.left), s.right = v("right", s.right), s.top = v("top", s.top), s.bottom = v("bottom", s.bottom);
      }
    }
    let o = "LEFT", h = "COLUMN";
    if (n) {
      const S = n.getElementsByTagNameNS(this.state.ns.w, "jc")[0];
      if (S) {
        const v = S.getAttribute("w:val") || "left";
        v === "center" ? (o = "CENTER", h = "COLUMN") : v === "right" ? (o = "RIGHT", h = "COLUMN") : (o = "LEFT", h = "COLUMN");
      }
      const R = n.getElementsByTagNameNS(this.state.ns.w, "tblpPr")[0];
      if (R) {
        const v = R.getAttribute("w:tblpXSpec") || "";
        v === "center" ? (o = "CENTER", h = "PAGE") : v === "right" && (o = "RIGHT", h = "PAGE");
      }
    }
    let l = null, f = null;
    if (a) {
      let S = a;
      for (; S; ) {
        const R = this.state.docStyles.get(S);
        if (!R) break;
        !i && R.tblBorders && (i = R.tblBorders), !l && R.tcBorders && (l = R.tcBorders), !f && R.shd && (f = R.shd), S = R.basedOn;
      }
    }
    const d = r.getElementsByTagNameNS(this.state.ns.w, "tblGrid")[0], u = d ? Array.from(d.getElementsByTagNameNS(this.state.ns.w, "gridCol")) : [], c = [];
    let g = 0;
    for (const S of u) {
      const v = parseInt(S.getAttribute("w:w") || "0") * 5;
      c.push(v), g += v;
    }
    const p = r.getElementsByTagNameNS(this.state.ns.w, "tr"), y = p.length, m = c.length > 0 ? c.length : 1, A = [];
    for (let S = 0; S < y; S++) {
      A[S] = [];
      for (let R = 0; R < m; R++)
        A[S][R] = { occupied: !1, rowSpan: 1, colSpan: 1, startRow: S, startCol: R, cell: null, vMerge: null };
    }
    for (let S = 0; S < y; S++) {
      const v = p[S].getElementsByTagNameNS(this.state.ns.w, "tc");
      let O = 0;
      for (let C = 0; C < v.length; C++) {
        const B = v[C];
        for (; O < m && A[S][O].occupied; ) O++;
        if (O >= m) break;
        const Z = B.getElementsByTagNameNS(this.state.ns.w, "tcPr")[0];
        let T = 1, j = null;
        if (Z) {
          const _ = Z.getElementsByTagNameNS(this.state.ns.w, "gridSpan")[0];
          _ && (T = parseInt(_.getAttribute("w:val") || "1"));
          const K = Z.getElementsByTagNameNS(this.state.ns.w, "vMerge")[0];
          K && (j = K.getAttribute("w:val") || "continue");
        }
        for (let _ = 0; _ < T; _++)
          O + _ < m && (A[S][O + _].occupied = !0, A[S][O + _].cell = B, _ === 0 && (A[S][O].colSpan = T, A[S][O].vMerge = j));
        O += T;
      }
    }
    for (let S = 0; S < m; S++)
      for (let R = 0; R < y; R++) {
        const v = A[R][S];
        if (v.cell && (v.vMerge === "restart" || !v.vMerge)) {
          let O = 1;
          for (let C = R + 1; C < y; C++) {
            const B = A[C][S];
            if (B.cell && B.vMerge === "continue")
              O++, B.isMergedBelow = !0;
            else break;
          }
          v.rowSpan = O;
        }
      }
    if (g === 0 && (g = this.state.TEXT_WIDTH, c.length === 0 && m > 0)) {
      const S = Math.floor(g / m);
      for (let R = 0; R < m; R++) c.push(S);
    }
    let E = 0;
    for (let S = 0; S < y; S++) {
      const R = p[S];
      let v = 0;
      const O = R.getElementsByTagNameNS(this.state.ns.w, "trPr")[0];
      if (O) {
        const C = O.getElementsByTagNameNS(this.state.ns.w, "trHeight")[0];
        C && (v = parseInt(C.getAttribute("w:val") || "0") * 5);
      }
      E += v > 0 ? v : 1500;
    }
    let x = `<hp:p id="${this.state.idCounter++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:tbl id="${e}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="1" rowCnt="${y}" colCnt="${m}" cellSpacing="0" borderFillIDRef="3" noAdjust="0">`;
    x += `<hp:sz width="${g}" widthRelTo="ABSOLUTE" height="${E}" heightRelTo="ABSOLUTE" protect="0"/>`, x += `<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="${h}" vertAlign="TOP" horzAlign="${o}" vertOffset="0" horzOffset="0"/>`, x += '<hp:outMargin left="0" right="0" top="0" bottom="0"/>', x += '<hp:inMargin left="540" right="540" top="0" bottom="0"/>';
    for (let S = 0; S < y; S++) {
      x += "<hp:tr>";
      let R = 0;
      for (; R < m; ) {
        const v = A[S][R];
        if (v.isMergedBelow) {
          R++;
          continue;
        }
        const O = v.cell;
        let C = 0;
        for (let se = 0; se < v.colSpan; se++) C += c[R + se] || 0;
        if (!O) {
          x += '<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="4">', x += '<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">', x += `<hp:p id="${this.state.idCounter++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t></hp:t></hp:run><hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850" spacing="0" horzpos="0" horzsize="${C}" flags="393216"/></hp:linesegarray></hp:p>`, x += "</hp:subList>", x += `<hp:cellAddr colAddr="${R}" rowAddr="${S}"/>`, x += `<hp:cellSpan colSpan="${v.colSpan}" rowSpan="${v.rowSpan}"/>`, x += `<hp:cellSz width="${C}" height="0"/>`, x += '<hp:cellMargin left="255" right="255" top="141" bottom="141"/>', x += "</hp:tc>", R++;
          continue;
        }
        const B = O.getElementsByTagNameNS(this.state.ns.w, "tcPr")[0], Z = this.getCellBorderFill(B || null, i, S, R, y, m, v.rowSpan, v.colSpan, l, f);
        let T = { ...s };
        if (B) {
          const se = B.getElementsByTagNameNS(this.state.ns.w, "tcMar")[0];
          if (se) {
            const P = (X) => {
              const de = se.getElementsByTagNameNS(this.state.ns.w, X)[0];
              return de ? parseInt(de.getAttribute("w:w") || "0") * 5 : null;
            }, G = P("left");
            G !== null && (T.left = G);
            const U = P("right");
            U !== null && (T.right = U);
            const Q = P("top");
            Q !== null && (T.top = Q);
            const M = P("bottom");
            M !== null && (T.bottom = M);
          }
        }
        const j = B && B.getElementsByTagNameNS(this.state.ns.w, "tcMar")[0] ? "1" : "0";
        let _ = "TOP";
        if (B) {
          const se = B.getElementsByTagNameNS(this.state.ns.w, "vAlign")[0];
          if (se) {
            const P = se.getAttribute("w:val");
            P === "center" ? _ = "CENTER" : P === "bottom" && (_ = "BOTTOM");
          }
        }
        x += `<hp:tc name="" header="0" hasMargin="${j}" protect="0" editable="0" dirty="0" borderFillIDRef="${Z.id}">`, x += `<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="${_}" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">`;
        let K = !1;
        for (let se = 0; se < O.childNodes.length; se++) {
          const P = O.childNodes[se];
          P.nodeName === "w:p" ? (x += this.convertParagraph(P), K = !0) : P.nodeName === "w:tbl" && (x += this.convertTable(P), K = !0);
        }
        K || (x += `<hp:p id="${this.state.idCounter++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t></hp:t></hp:run><hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850" spacing="0" horzpos="0" horzsize="${C}" flags="393216"/></hp:linesegarray></hp:p>`), x += "</hp:subList>", x += `<hp:cellAddr colAddr="${R}" rowAddr="${S}"/>`, x += `<hp:cellSpan colSpan="${v.colSpan}" rowSpan="${v.rowSpan}"/>`, x += `<hp:cellSz width="${C}" height="0"/>`, x += `<hp:cellMargin left="${T.left}" right="${T.right}" top="${T.top}" bottom="${T.bottom}"/>`, x += "</hp:tc>", R += v.colSpan;
      }
      x += "</hp:tr>";
    }
    return x += "</hp:tbl></hp:run>", x += `<hp:linesegarray><hp:lineseg textpos="0" vertpos="${this.state.currentVertPos}" vertsize="${E}" textheight="${E}" baseline="${E}" spacing="0" horzpos="0" horzsize="${g}" flags="393216"/></hp:linesegarray>`, x += "</hp:p>", this.state.currentVertPos += E, x;
  }
  convertDrawing(r) {
    var o;
    const e = r.getElementsByTagName("*"), n = Array.from(e).filter((h) => h.localName === "blip");
    if (console.log(`✅ 찾은 이미지(blip) 개수: ${n.length}`), n.length > 0) {
      const h = n[0].getAttributeNS(this.state.ns.r, "embed") || n[0].getAttribute("r:embed") || n[0].getAttribute("embed");
      return console.log(`- 추출된 이미지 연결 ID: ${h}`), this.convertPicture(r, n[0]);
    }
    const i = r.getElementsByTagNameNS("http://schemas.microsoft.com/office/word/2010/wordprocessingShape", "txbx");
    let a = "";
    if (i.length > 0) {
      const h = i[0].getElementsByTagNameNS(this.state.ns.w, "p");
      for (const l of h) {
        const f = l.getElementsByTagNameNS(this.state.ns.w, "t");
        for (const d of f) a += d.textContent + `
`;
      }
    }
    return a ? `<hp:run charPrIDRef="0"><hp:t>[텍스트 상자: ${Mt(a.trim())}]</hp:t></hp:run>` : r.getElementsByTagName("c:chart").length > 0 || (o = r.textContent) != null && o.includes("chart") ? '<hp:run charPrIDRef="0"><hp:t>[차트: 변환되지 않음]</hp:t></hp:run>' : "";
  }
  convertPicture(r, e) {
    const n = e.getAttributeNS(this.state.ns.r, "embed") || e.getAttribute("r:embed") || e.getAttribute("embed");
    if (!n) return "";
    const i = this.state.images.get(n);
    if (console.log(this.state.images), !i) return "";
    const a = this.state.tableIdCounter++;
    let s = 27577, o = 19913;
    const h = r.getElementsByTagNameNS(this.state.ns.wp, "extent");
    if (h.length > 0) {
      const C = h[0].getAttribute("cx"), B = h[0].getAttribute("cy");
      C && B && (s = Math.round(parseInt(C) * this.state.HWPUNIT_PER_INCH / this.state.EMU_PER_INCH), o = Math.round(parseInt(B) * this.state.HWPUNIT_PER_INCH / this.state.EMU_PER_INCH));
    }
    const l = Math.floor(Math.random() * 2e9), f = r.getElementsByTagNameNS(this.state.ns.wp, "inline").length > 0, d = r.getElementsByTagNameNS(this.state.ns.wp, "anchor")[0];
    let u = f ? "1" : "0", c = "TOP_AND_BOTTOM", g = "PARA", p = "COLUMN", y = "TOP", m = "LEFT", A = "0", E = "0", x = "0", S = "0", R = "0", v = "0";
    if (d) {
      const C = d.getElementsByTagNameNS(this.state.ns.wp, "wrapSquare")[0], B = d.getElementsByTagNameNS(this.state.ns.wp, "wrapTight")[0], Z = d.getElementsByTagNameNS(this.state.ns.wp, "wrapThrough")[0], T = d.getElementsByTagNameNS(this.state.ns.wp, "wrapNone")[0], j = d.getElementsByTagNameNS(this.state.ns.wp, "wrapTopAndBottom")[0];
      C ? c = "SQUARE" : B || Z ? c = "TIGHT" : T ? c = "BEHIND_TEXT" : j && (c = "TOP_AND_BOTTOM");
      const _ = d.getElementsByTagNameNS(this.state.ns.wp, "positionV")[0];
      if (_) {
        const P = _.getAttribute("relativeFrom");
        P === "page" ? g = "PAGE" : P === "paragraph" ? g = "PARA" : P === "margin" && (g = "PAPER");
        const G = _.getElementsByTagNameNS(this.state.ns.wp, "posOffset")[0];
        G && (A = String(Math.round(parseInt(G.textContent || "0") * this.state.HWPUNIT_PER_INCH / this.state.EMU_PER_INCH)));
        const U = _.getElementsByTagNameNS(this.state.ns.wp, "align")[0];
        if (U) {
          const Q = U.textContent;
          Q === "center" ? y = "CENTER" : Q === "bottom" ? y = "BOTTOM" : Q === "top" && (y = "TOP");
        }
      }
      const K = d.getElementsByTagNameNS(this.state.ns.wp, "positionH")[0];
      if (K) {
        const P = K.getAttribute("relativeFrom");
        P === "page" ? p = "PAGE" : P === "column" ? p = "COLUMN" : P === "margin" && (p = "PAPER");
        const G = K.getElementsByTagNameNS(this.state.ns.wp, "posOffset")[0];
        G && (E = String(Math.round(parseInt(G.textContent || "0") * this.state.HWPUNIT_PER_INCH / this.state.EMU_PER_INCH)));
        const U = K.getElementsByTagNameNS(this.state.ns.wp, "align")[0];
        if (U) {
          const Q = U.textContent;
          Q === "center" ? m = "CENTER" : Q === "right" ? m = "RIGHT" : Q === "left" && (m = "LEFT");
        }
      }
      const se = d.getElementsByTagNameNS(this.state.ns.wp, "effectExtent")[0];
      if (se) {
        const P = (G) => String(Math.round(parseInt(G || "0") * this.state.HWPUNIT_PER_INCH / this.state.EMU_PER_INCH));
        x = P(se.getAttribute("l")), S = P(se.getAttribute("r")), R = P(se.getAttribute("t")), v = P(se.getAttribute("b"));
      }
    }
    let O = '<hp:run charPrIDRef="0">';
    return O += `<hp:pic id="${a}" zOrder="0" numberingType="NONE" textWrap="${c}" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" href="" groupLevel="0" instId="${l}" reverse="0">`, O += '<hp:offset x="0" y="0"/>', O += `<hp:orgSz width="${s}" height="${o}"/>`, O += '<hp:curSz width="0" height="0"/>', O += '<hp:flip horizontal="0" vertical="0"/>', O += `<hp:rotationInfo angle="0" centerX="${Math.round(s / 2)}" centerY="${Math.round(o / 2)}" rotateimage="1"/>`, O += "<hp:renderingInfo>", O += '<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>', O += '<hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>', O += '<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>', O += "</hp:renderingInfo>", O += "<hp:imgRect>", O += '<hc:pt0 x="0" y="0"/>', O += `<hc:pt1 x="${s}" y="0"/>`, O += `<hc:pt2 x="${s}" y="${o}"/>`, O += `<hc:pt3 x="0" y="${o}"/>`, O += "</hp:imgRect>", O += '<hp:imgClip left="0" right="0" top="0" bottom="0"/>', O += '<hp:inMargin left="0" right="0" top="0" bottom="0"/>', O += `<hc:img binaryItemIDRef="${i.manifestId}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/>`, O += "<hp:effects/>", O += `<hp:sz width="${s}" widthRelTo="ABSOLUTE" height="${o}" heightRelTo="ABSOLUTE" protect="0"/>`, O += `<hp:pos treatAsChar="${u}" affectLSpacing="0" flowWithText="1" allowOverlap="1" holdAnchorAndSO="0" vertRelTo="${g}" horzRelTo="${p}" vertAlign="${y}" horzAlign="${m}" vertOffset="${A}" horzOffset="${E}"/>`, O += `<hp:outMargin left="${x}" right="${S}" top="${R}" bottom="${v}"/>`, O += `<hp:shapeComment>그림입니다.
원본 그림의 이름: ${i.manifestId}.${i.ext}
원본 그림의 크기: 가로 ${Math.round(s / 283.465)}mm, 세로 ${Math.round(o / 283.465)}mm</hp:shapeComment>`, O += "</hp:pic></hp:run>", O;
  }
  createParaProperties() {
    let r = `<hh:paraProperties itemCnt="${this.state.paraShapes.size}">`;
    for (const [e, n] of this.state.paraShapes)
      r += `<hh:paraPr id="${e}" tabPrIDRef="${n.tabPrIDRef || "1"}" condense="0" fontLineHeight="0" snapToGrid="1" suppressLineNumbers="0" checked="0">`, r += `<hh:align horizontal="${n.align}" vertical="BASELINE"/>`, r += `<hh:heading type="${n.heading}" idRef="${n.headingIdRef || "0"}" level="${n.level || "0"}"/>`, r += `<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="BREAK_WORD" widowOrphan="0" keepWithNext="${n.keepWithNext || "0"}" keepLines="${n.keepLines || "0"}" pageBreakBefore="${n.pageBreakBefore || "0"}" lineWrap="BREAK"/>`, r += '<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>', r += `<hh:margin><hc:intent value="${n.indent || 0}" unit="HWPUNIT"/><hc:left value="${n.leftMargin || 0}" unit="HWPUNIT"/><hc:right value="${n.rightMargin || 0}" unit="HWPUNIT"/><hc:prev value="${n.prevSpacing || 0}" unit="HWPUNIT"/><hc:next value="${n.nextSpacing || 0}" unit="HWPUNIT"/></hh:margin>`, r += `<hh:lineSpacing type="${n.lineSpacingType}" value="${n.lineSpacingVal}" unit="HWPUNIT"/>`, r += `<hh:border borderFillIDRef="${n.borderFillIDRef}" offsetLeft="400" offsetRight="400" offsetTop="100" offsetBottom="100" connect="0" ignoreMargin="0"/>`, r += "</hh:paraPr>";
    return r += "</hh:paraProperties>", r;
  }
  createCharProperties() {
    let r = `<hh:charProperties itemCnt="${this.state.charShapes.size}">`;
    for (const [e, n] of this.state.charShapes) {
      r += `<hh:charPr id="${e}" height="${n.height}" textColor="${n.textColor}" shadeColor="${n.shadeColor}" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="${n.borderFillIDRef}">`;
      const i = n.fontIds || { hangulId: 0, latinId: 0, hanjaId: 0, japaneseId: 0, otherId: 0, symbolId: 0, userId: 0 };
      r += `<hh:fontRef hangul="${i.hangulId}" latin="${i.latinId}" hanja="${i.hanjaId}" japanese="${i.japaneseId}" other="${i.otherId}" symbol="${i.symbolId}" user="${i.userId}"/>`, r += '<hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>', r += '<hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>', r += '<hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>', r += '<hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>', n.bold && (r += "<hh:bold/>"), n.italic && (r += "<hh:italic/>"), n.underline && (r += `<hh:underline type="BOTTOM" shape="SOLID" color="${n.textColor}"/>`), n.strike && (r += `<hh:strikeout shape="SOLID" color="${n.textColor}"/>`), n.supscript && (r += "<hh:supscript/>"), n.subscript && (r += "<hh:subscript/>"), r += "</hh:charPr>";
    }
    return r += "</hh:charProperties>", r;
  }
  createContentHpf() {
    let r = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
`;
    r += `<opf:package xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:opf="http://www.idpf.org/2007/opf/" xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" xmlns:epub="http://www.idpf.org/2007/ops" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" version="" unique-identifier="" id="">
`, r += `<opf:metadata>
`, r += `  <opf:title>문서</opf:title>
`, r += `  <opf:language>ko</opf:language>
`, r += `  <opf:meta name="subject" content="text"></opf:meta>
`, r += `</opf:metadata>
`, r += `<opf:manifest>
`;
    for (const [, e] of this.state.images) {
      const n = e.ext === "png" ? "image/png" : e.ext === "jpg" || e.ext === "jpeg" ? "image/jpeg" : e.ext === "gif" ? "image/gif" : "image/png";
      r += `  <opf:item id="${e.manifestId}" href="BinData/${e.manifestId}.${e.ext}" media-type="${n}" isEmbeded="1"/>
`;
    }
    return r += `  <opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>
`, r += `  <opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>
`, r += `  <opf:item id="settings" href="settings.xml" media-type="application/xml"/>
`, r += `</opf:manifest>
`, r += `<opf:spine>
`, r += `  <opf:itemref idref="header"/>
`, r += `  <opf:itemref idref="section0"/>
`, r += `</opf:spine>
`, r += "</opf:package>", r;
  }
  createContainer() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"><ocf:rootfiles><ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/><ocf:rootfile full-path="Preview/PrvText.txt" media-type="text/plain"/><ocf:rootfile full-path="META-INF/container.rdf" media-type="application/rdf+xml"/></ocf:rootfiles></ocf:container>';
  }
  createContainerhdf() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"><rdf:Description rdf:about=""><ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#" rdf:resource="Contents/header.xml"/></rdf:Description><rdf:Description rdf:about="Contents/header.xml"><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#HeaderFile"/></rdf:Description><rdf:Description rdf:about=""><ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#" rdf:resource="Contents/section0.xml"/></rdf:Description><rdf:Description rdf:about="Contents/section0.xml"><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#SectionFile"/></rdf:Description><rdf:Description rdf:about=""><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#Document"/></rdf:Description></rdf:RDF>';
  }
  createManifest() {
    let r = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
`;
    r += `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
`, r += `  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/hwpml-package+xml"/>
`, r += `  <manifest:file-entry manifest:full-path="version.xml" manifest:media-type="application/xml"/>
`, r += `  <manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="application/xml"/>
`, r += `  <manifest:file-entry manifest:full-path="Contents/content.hpf" manifest:media-type="application/hwpml-package+xml"/>
`, r += `  <manifest:file-entry manifest:full-path="Contents/header.xml" manifest:media-type="application/xml"/>
`, r += `  <manifest:file-entry manifest:full-path="Contents/section0.xml" manifest:media-type="application/xml"/>
`;
    for (const [, e] of this.state.images) {
      const n = e.ext === "png" ? "image/png" : e.ext === "jpg" || e.ext === "jpeg" ? "image/jpeg" : e.ext === "gif" ? "image/gif" : "image/png";
      r += `  <manifest:file-entry manifest:full-path="Contents/${e.path}" manifest:media-type="${n}"/>
`;
    }
    return r += "</manifest:manifest>", r;
  }
  createHwpx(r) {
    const e = this.createSection(r), n = this.createHeader();
    return {
      mimetype: "application/hwp+zip",
      version: this.createVersion(),
      settings: this.createSettings(),
      header: n,
      section0: e,
      contentHpf: this.createContentHpf(),
      container: this.createContainer(),
      containerRdf: this.createContainerhdf(),
      manifest: this.createManifest()
    };
  }
  createVersion() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hv:HCFVersion xmlns:hv="${this.state.hwpxNs.hv}" tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="0" buildNumber="1" os="6" xmlVersion="1.2" application="Hancom Office Hangul" appVersion="11, 20, 0, 1520"/>`;
  }
  createSettings() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ha:HWPApplicationSetting xmlns:ha="${this.state.hwpxNs.ha}" xmlns:config="${this.state.hwpxNs.config}"><ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/></ha:HWPApplicationSetting>`;
  }
  createHeader() {
    let r = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hh:head xmlns:ha="${this.state.hwpxNs.ha}" xmlns:hp="${this.state.hwpxNs.hp}" xmlns:hp10="${this.state.hwpxNs.hp10}" xmlns:hs="${this.state.hwpxNs.hs}" xmlns:hc="${this.state.hwpxNs.hc}" xmlns:hh="${this.state.hwpxNs.hh}" xmlns:hhs="${this.state.hwpxNs.hhs}" xmlns:hm="${this.state.hwpxNs.hm}" xmlns:hpf="${this.state.hwpxNs.hpf}" xmlns:dc="${this.state.hwpxNs.dc}" xmlns:opf="${this.state.hwpxNs.opf}" xmlns:ooxmlchart="${this.state.hwpxNs.ooxmlchart}" xmlns:epub="${this.state.hwpxNs.epub}" xmlns:config="${this.state.hwpxNs.config}" version="1.2" secCnt="1">`;
    return r += '<hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>', r += "<hh:refList>", r += this.createFontFaces(), r += this.createBorderFills(), r += this.createBinDataProperties(), r += this.createCharProperties(), r += this.createTabProperties(), r += this.createNumberings(), r += this.createParaProperties(), r += this.createStyles(), r += '<hh:memoProperties itemCnt="1"><hh:memoPr id="1" width="15024" lineWidth="1" lineType="SOLID" lineColor="#000000" fillColor="#CCFF99" activeColor="#FFFF99" memoType="NOMAL"/></hh:memoProperties>', r += "</hh:refList>", r += `    <hh:compatibleDocument targetProgram="MS_WORD">
        <hh:layoutCompatibility>
            <hh:applyFontWeightToBold />
            <hh:useInnerUnderline />
            <hh:useLowercaseStrikeout />
            <hh:extendLineheightToOffset />
            <hh:treatQuotationAsLatin />
            <hh:doNotAlignWhitespaceOnRight />
            <hh:doNotAdjustWordInJustify />
            <hh:baseCharUnitOnEAsian />
            <hh:baseCharUnitOfIndentOnFirstChar />
            <hh:adjustLineheightToFont />
            <hh:adjustBaselineInFixedLinespacing />
            <hh:applyPrevspacingBeneathObject />
            <hh:applyNextspacingOfLastPara />
            <hh:adjustParaBorderfillToSpacing />
            <hh:connectParaBorderfillOfEqualBorder />
            <hh:adjustParaBorderOffsetWithBorder />
            <hh:extendLineheightToParaBorderOffset />
            <hh:applyParaBorderToOutside />
            <hh:applyMinColumnWidthTo1mm />
            <hh:applyTabPosBasedOnSegment />
            <hh:breakTabOverLine />
            <hh:adjustVertPosOfLine />
            <hh:doNotAlignLastForbidden />
            <hh:adjustMarginFromAdjustLineheight />
            <hh:baseLineSpacingOnLineGrid />
            <hh:applyCharSpacingToCharGrid />
            <hh:doNotApplyGridInHeaderFooter />
            <hh:applyExtendHeaderFooterEachSection />
            <hh:doNotApplyLinegridAtNoLinespacing />
            <hh:doNotAdjustEmptyAnchorLine />
            <hh:overlapBothAllowOverlap />
            <hh:extendVertLimitToPageMargins />
            <hh:doNotHoldAnchorOfTable />
            <hh:doNotFormattingAtBeneathAnchor />
            <hh:adjustBaselineOfObjectToBottom />
        </hh:layoutCompatibility>
    </hh:compatibleDocument>`, r += '<hh:docOption><hh:linkinfo path="" pageInherit="0" footnoteInherit="0"/></hh:docOption>', r += '<hh:trackchageConfig flags="0"/>', r += "</hh:head>", r;
  }
  createBorderFills() {
    let r = `<hh:borderFills itemCnt="${this.state.borderFills.size}">`;
    for (const [e, n] of this.state.borderFills) {
      const i = n.backColor || "#FFFFFF", a = n.slash || { type: "NONE" }, s = n.backSlash || { type: "NONE" };
      r += `<hh:borderFill id="${e}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">`, r += `<hh:slash type="${a.type}" Crooked="0" isCounter="0" />`, r += `<hh:backSlash type="${s.type}" Crooked="0" isCounter="0" />`, r += `<hh:leftBorder type="${n.leftBorder.type}" width="${n.leftBorder.width}" color="${n.leftBorder.color}" />`, r += `<hh:rightBorder type="${n.rightBorder.type}" width="${n.rightBorder.width}" color="${n.rightBorder.color}" />`, r += `<hh:topBorder type="${n.topBorder.type}" width="${n.topBorder.width}" color="${n.topBorder.color}" />`, r += `<hh:bottomBorder type="${n.bottomBorder.type}" width="${n.bottomBorder.width}" color="${n.bottomBorder.color}" />`, r += '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000" />', r += `<hc:fillBrush><hc:winBrush faceColor="${i}" hatchColor="#000000" alpha="0" /></hc:fillBrush>`, r += "</hh:borderFill>";
    }
    return r += "</hh:borderFills>", r;
  }
  createBinDataProperties() {
    if (this.state.images.size === 0) return "";
    let r = `<hh:binDataProperties itemCnt="${this.state.images.size}">`;
    for (const [, e] of this.state.images)
      r += `<hh:binData id="${e.binDataId}" itemIDRef="${e.manifestId}" extension="${e.ext}" format="${e.ext}" type="EMBEDDING" state="NOT_ACCESSED"/>`;
    return r += "</hh:binDataProperties>", r;
  }
  createTabProperties() {
    let r = '<hh:tabProperties itemCnt="2">';
    r += '<hh:tabPr id="0" itemCnt="0" autoTabLeft="1" autoTabRight="0"/>', r += '<hh:tabPr id="1" itemCnt="2" autoTabLeft="0" autoTabRight="0">';
    const e = Math.round((this.state.PAGE_WIDTH - this.state.MARGIN_LEFT - this.state.MARGIN_RIGHT) / 2), n = this.state.PAGE_WIDTH - this.state.MARGIN_LEFT - this.state.MARGIN_RIGHT;
    return r += `<hh:tabItem pos="${e}" type="CENTER" leader="NONE"/>`, r += `<hh:tabItem pos="${n}" type="RIGHT" leader="NONE"/>`, r += "</hh:tabPr>", r += "</hh:tabProperties>", r;
  }
  createNumberings() {
    var a;
    if (this.state.abstractNumMap.size === 0)
      return '<hh:numberings itemCnt="0"/>';
    const r = {
      decimal: "DIGIT",
      lowerRoman: "ROMAN_SMALL",
      upperRoman: "ROMAN_CAPITAL",
      lowerLetter: "LATIN_SMALL",
      upperLetter: "LATIN_CAPITAL",
      bullet: "BULLET",
      ganada: "DIGIT",
      chosung: "DIGIT",
      none: "DIGIT"
    }, e = {
      "": "·",
      o: "o",
      "§": "§"
    };
    let n = `<hh:numberings itemCnt="${this.state.abstractNumMap.size}">`, i = 1;
    for (const [s, o] of this.state.abstractNumMap) {
      n += `<hh:numbering id="${i}" start="${((a = o.get("0")) == null ? void 0 : a.start) || 1}">`;
      for (let h = 1; h <= 10; h++) {
        const l = o.get(String(h - 1));
        if (l) {
          const f = r[l.numFmt] || "DIGIT", d = l.start || 1;
          if (l.numFmt === "bullet") {
            const u = e[l.lvlText] || "·";
            n += `<hh:paraHead start="${d}" level="${h}" align="LEFT" useInstWidth="1" autoIndent="1"`, n += ' widthAdjust="0" textOffsetType="PERCENT" textOffset="100" numFormat="BULLET"', n += ` charPrIDRef="0" checkable="0">${u}</hh:paraHead>`;
          } else {
            let u = this.convertLvlTextToHwpx(l.lvlText, f, h);
            n += `<hh:paraHead start="${d}" level="${h}" align="LEFT" useInstWidth="1" autoIndent="1"`, n += ` widthAdjust="0" textOffsetType="PERCENT" textOffset="100" numFormat="${f}"`, n += ` charPrIDRef="0" checkable="0">${Mt(u)}</hh:paraHead>`;
          }
        } else
          n += `<hh:paraHead start="1" level="${h}" align="LEFT" useInstWidth="1" autoIndent="0"`, n += ' widthAdjust="0" textOffsetType="PERCENT" textOffset="100" numFormat="DIGIT"', n += ' charPrIDRef="0" checkable="0">^N</hh:paraHead>';
      }
      n += "</hh:numbering>", i++;
    }
    return n += "</hh:numberings>", n;
  }
  convertLvlTextToHwpx(r, e, n) {
    if (!r) return "^N.";
    let i = r;
    const a = `%${n}`;
    i = i.replace(a, "^N");
    for (let s = 1; s <= 10; s++)
      s !== n && (i = i.replace(`%${s}`, ""));
    return i;
  }
  createStyles() {
    const r = [
      { id: 0, type: "PARA", name: "바탕글", engName: "Normal", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 0 },
      { id: 1, type: "CHAR", name: "Book Title", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 1 },
      { id: 2, type: "PARA", name: "Caption", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 2 },
      { id: 3, type: "CHAR", name: "Default Paragraph Font", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 3 },
      { id: 4, type: "CHAR", name: "Emphasis", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 4 },
      { id: 5, type: "CHAR", name: "Endnote Text Char", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 5 },
      { id: 6, type: "CHAR", name: "FollowedHyperlink", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 6 },
      { id: 7, type: "PARA", name: "Footer", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 7 },
      { id: 8, type: "CHAR", name: "Footer Char", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 8 },
      { id: 9, type: "CHAR", name: "Footnote Text Char", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 9 },
      { id: 10, type: "PARA", name: "Header", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 10 },
      { id: 11, type: "CHAR", name: "Header Char", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 11 },
      { id: 12, type: "PARA", name: "Heading 1", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 12 },
      { id: 13, type: "CHAR", name: "Heading 1 Char", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 13 },
      { id: 14, type: "PARA", name: "Heading 2", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 14 },
      { id: 15, type: "CHAR", name: "Heading 2 Char", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 15 },
      { id: 16, type: "PARA", name: "Heading 3", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 16 },
      { id: 17, type: "CHAR", name: "Heading 3 Char", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 17 },
      { id: 18, type: "PARA", name: "Heading 4", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 18 },
      { id: 19, type: "CHAR", name: "Heading 4 Char", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 19 },
      { id: 20, type: "PARA", name: "Heading 5", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 20 },
      { id: 21, type: "CHAR", name: "Heading 5 Char", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 21 },
      { id: 22, type: "PARA", name: "Heading 6", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 22 },
      { id: 23, type: "CHAR", name: "Heading 6 Char", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 23 },
      { id: 24, type: "PARA", name: "Heading 7", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 24 },
      { id: 25, type: "CHAR", name: "Heading 7 Char", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 25 },
      { id: 26, type: "PARA", name: "Heading 8", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 26 },
      { id: 27, type: "CHAR", name: "Heading 8 Char", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 27 },
      { id: 28, type: "PARA", name: "Heading 9", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 28 },
      { id: 29, type: "CHAR", name: "Heading 9 Char", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 29 },
      { id: 30, type: "CHAR", name: "Hyperlink", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 30 },
      { id: 31, type: "CHAR", name: "Intense Emphasis", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 31 },
      { id: 32, type: "PARA", name: "Intense Quote", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 32 },
      { id: 33, type: "CHAR", name: "Intense Quote Char", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 33 },
      { id: 34, type: "CHAR", name: "Intense Reference", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 34 },
      { id: 35, type: "PARA", name: "List Paragraph", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 35 },
      { id: 36, type: "PARA", name: "No Spacing", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 36 },
      { id: 37, type: "CHAR", name: "Placeholder Text", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 37 },
      { id: 38, type: "PARA", name: "Quote", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 38 },
      { id: 39, type: "CHAR", name: "Quote Char", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 39 },
      { id: 40, type: "CHAR", name: "Strong", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 40 },
      { id: 41, type: "PARA", name: "Subtitle", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 41 },
      { id: 42, type: "CHAR", name: "Subtitle Char", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 42 },
      { id: 43, type: "CHAR", name: "Subtle Emphasis", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 43 },
      { id: 44, type: "CHAR", name: "Subtle Reference", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 44 },
      { id: 45, type: "PARA", name: "TOC Heading", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 45 },
      { id: 46, type: "PARA", name: "Title", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 46 },
      { id: 47, type: "CHAR", name: "Title Char", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 47 },
      { id: 48, type: "CHAR", name: "endnote reference", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 48 },
      { id: 49, type: "PARA", name: "endnote text", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 49 },
      { id: 50, type: "CHAR", name: "footnote reference", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 50 },
      { id: 51, type: "PARA", name: "footnote text", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 51 },
      { id: 52, type: "PARA", name: "table of figures", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 52 },
      { id: 53, type: "PARA", name: "toc 1", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 53 },
      { id: 54, type: "PARA", name: "toc 2", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 54 },
      { id: 55, type: "PARA", name: "toc 3", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 55 },
      { id: 56, type: "PARA", name: "toc 4", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 56 },
      { id: 57, type: "PARA", name: "toc 5", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 57 },
      { id: 58, type: "PARA", name: "toc 6", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 58 },
      { id: 59, type: "PARA", name: "toc 7", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 59 },
      { id: 60, type: "PARA", name: "toc 8", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 60 },
      { id: 61, type: "PARA", name: "toc 9", engName: "", paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 61 }
    ];
    this.state.docxStyleToHwpxId = {
      Normal: 0,
      Heading1: 12,
      Heading2: 14,
      Heading3: 16,
      Heading4: 18,
      Heading5: 20,
      Heading6: 22,
      Heading7: 24,
      Heading8: 26,
      Heading9: 28,
      Title: 46,
      Subtitle: 41,
      ListParagraph: 35,
      NoSpacing: 36,
      Quote: 38,
      IntenseQuote: 32,
      Header: 10,
      Footer: 7,
      FootnoteText: 51,
      EndnoteText: 49,
      TOCHeading: 45,
      TOC1: 53,
      TOC2: 54,
      TOC3: 55,
      TOC4: 56,
      TOC5: 57,
      TOC6: 58,
      TOC7: 59,
      TOC8: 60,
      TOC9: 61,
      Caption: 2,
      TableofFigures: 52
    };
    let e = `<hh:styles itemCnt="${r.length}">`;
    for (const n of r)
      e += `<hh:style id="${n.id}" type="${n.type}" name="${Mt(n.name)}" engName="${Mt(n.engName)}" paraPrIDRef="${n.paraPrIDRef}" charPrIDRef="${n.charPrIDRef}" nextStyleIDRef="${n.nextStyleIDRef}" langID="1042" lockForm="0"/>`;
    return e += "</hh:styles>", e;
  }
}
class Pi {
  constructor(r) {
    ce(this, "state");
    this.state = r;
  }
  async packageHwpx(r) {
    const e = new yt();
    e.file("mimetype", r.mimetype, { compression: "STORE" }), e.file("version.xml", r.version), e.file("settings.xml", r.settings), e.file("Contents/header.xml", r.header), e.file("Contents/section0.xml", r.section0), e.file("Contents/content.hpf", r.contentHpf), e.file("META-INF/container.xml", r.container), e.file("META-INF/container.rdf", r.containerRdf), e.file("META-INF/manifest.xml", r.manifest), e.file("Preview/PrvText.txt", "문서 미리보기");
    for (const [, n] of this.state.images)
      e.file(n.path, n.data);
    return console.log(`✅ Packaged HWPX with ${this.state.images.size} images`), await e.generateAsync({ type: "blob" });
  }
}
class Fi {
  async convert(r) {
    var u;
    console.log("[1/8] DOCX 로드 중...");
    const e = await yt.loadAsync(r), n = new Ei(), i = new Ri(n), a = new Ci(n), s = new Pi(n);
    console.log("[2/8] 메타데이터 파싱..."), await i.parseMetadata(e), console.log("[3/8] 이미지 및 관계 파싱..."), await i.parseRelationships(e), await i.parseImages(e), await i.parseStyles(e), await i.parseNumbering(e), await i.parseFootnotes(e), await i.parseEndnotes(e), console.log("[4/8] DOCX 파싱 중...");
    const o = await ((u = e.file("word/document.xml")) == null ? void 0 : u.async("text"));
    if (!o) throw new Error("word/document.xml not found");
    const l = new DOMParser().parseFromString(o, "text/xml");
    i.parsePageSetup(l), console.log("[5/8] 스타일 초기화..."), console.log("[6/8] HWPX 생성 중...");
    const f = a.createHwpx(l);
    console.log("[7/8] 패키징 중...");
    const d = await s.packageHwpx(f);
    return console.log("[8/8] 완료!"), d;
  }
}
const ki = async (t) => await new Fi().convert(t);
class Li {
  constructor() {
    ce(this, "ns", {
      hp: "http://www.hancom.co.kr/hwpml/2011/paragraph",
      hh: "http://www.hancom.co.kr/hwpml/2011/head",
      hc: "http://www.hancom.co.kr/hwpml/2011/core",
      hs: "http://www.hancom.co.kr/hwpml/2011/section",
      ha: "http://www.hancom.co.kr/hwpml/2011/app",
      opf: "http://www.idpf.org/2007/opf/"
    });
    ce(this, "charProperties", /* @__PURE__ */ new Map());
    ce(this, "paraProperties", /* @__PURE__ */ new Map());
    ce(this, "borderFills", /* @__PURE__ */ new Map());
    ce(this, "fontFaces", {});
    ce(this, "images", /* @__PURE__ */ new Map());
    ce(this, "imageRels", []);
    ce(this, "relIdCounter", 1);
    // 페이지 설정
    ce(this, "pageWidth", 11906);
    ce(this, "pageHeight", 16838);
    ce(this, "marginTop", 1440);
    ce(this, "marginBottom", 1440);
    ce(this, "marginLeft", 1800);
    ce(this, "marginRight", 1800);
    ce(this, "marginHeader", 851);
    ce(this, "marginFooter", 851);
  }
}
class Bi {
  constructor(r) {
    ce(this, "ctx");
    this.ctx = r;
  }
  async parseContentHpf(r) {
  }
  parseHeader(r) {
    const n = new DOMParser().parseFromString(r, "text/xml");
    this.parseFontFaces(n), this.parseBorderFills(n), this.parseCharProperties(n), this.parseParaProperties(n), this.parseStyles(n), this.parseNumberings(n), console.log(`header.xml 파싱 완료: charPr=${this.ctx.charProperties.size}, paraPr=${this.ctx.paraProperties.size}, borderFill=${this.ctx.borderFills.size}`);
  }
  parseFontFaces(r) {
    const e = r.getElementsByTagNameNS(this.ctx.ns.hh, "fontface");
    for (const n of Array.from(e)) {
      const i = n.getAttribute("lang") || "UNKNOWN", a = [], s = n.getElementsByTagNameNS(this.ctx.ns.hh, "font");
      for (const o of Array.from(s))
        a.push({
          id: o.getAttribute("id") || "",
          face: o.getAttribute("face") || ""
        });
      this.ctx.fontFaces[i] = a;
    }
  }
  parseBorderFills(r) {
    let e = [];
    const n = r.getElementsByTagName("*");
    for (let i = 0; i < n.length; i++)
      n[i].localName === "borderFill" && e.push(n[i]);
    for (const i of e) {
      const a = i.getAttribute("id") || "", s = (l) => {
        let f = i.getElementsByTagNameNS(this.ctx.ns.hh, l)[0];
        if (!f) {
          const d = i.getElementsByTagName(`hh:${l}`);
          d.length > 0 && (f = d[0]);
        }
        return f ? {
          type: f.getAttribute("type") || "none",
          width: f.getAttribute("width") || "0",
          color: f.getAttribute("color") || "#000000"
        } : { type: "none", width: "0", color: "#000000" };
      }, o = {
        id: a,
        leftBorder: s("leftBorder"),
        rightBorder: s("rightBorder"),
        topBorder: s("topBorder"),
        bottomBorder: s("bottomBorder")
      }, h = Array.from(i.childNodes).find((l) => l.nodeType === 1 && l.localName === "fillBrush");
      if (h) {
        const l = Array.from(h.childNodes).find((f) => f.nodeType === 1 && f.localName === "winBrush");
        if (l) {
          const f = l.getAttribute("faceColor");
          f && f !== "none" && (o.faceColor = f);
        }
      }
      this.ctx.borderFills.set(a, o);
    }
  }
  parseCharProperties(r) {
    var n, i, a, s;
    const e = r.getElementsByTagName("hh:charPr");
    for (let o = 0; o < e.length; o++) {
      const h = e[o], l = h.getAttribute("id") || "", f = parseInt(h.getAttribute("height") || "1000"), d = h.getAttribute("textColor") || "#000000", u = h.getAttribute("shadeColor") || "none", c = h.getAttribute("borderFillIDRef") || "", g = (B) => Array.from(h.childNodes).find(
        (Z) => Z.nodeType === 1 && Z.localName === B
      ), p = ((n = g("bold")) == null ? void 0 : n.getAttribute("value")) === "true", y = ((i = g("italic")) == null ? void 0 : i.getAttribute("value")) === "true", m = g("underline"), A = !!m && (m.getAttribute("type") || "NONE") !== "NONE", E = g("strikeout"), x = !!E && (E.getAttribute("type") || "NONE") !== "NONE", S = ((a = g("superScript")) == null ? void 0 : a.getAttribute("value")) === "true", R = ((s = g("subScript")) == null ? void 0 : s.getAttribute("value")) === "true";
      let v = "함초롬바탕", O = "함초롬바탕";
      const C = g("fontRef");
      if (C) {
        const B = C.getAttribute("hangul") || "0", Z = C.getAttribute("latin") || "0", T = this.ctx.fontFaces.HANGUL || this.ctx.fontFaces.Hangul || [], j = this.ctx.fontFaces.LATIN || this.ctx.fontFaces.Latin || [], _ = T.find((se) => se.id === B), K = j.find((se) => se.id === Z);
        _ != null && _.face && (v = _.face), K != null && K.face && (O = K.face);
      }
      this.ctx.charProperties.set(l, {
        id: l,
        height: f,
        textColor: d,
        shadeColor: u,
        borderFillIDRef: c,
        bold: p,
        italic: y,
        underline: A,
        strikeout: x,
        supscript: S,
        subscript: R,
        hangulFont: v,
        latinFont: O
      });
    }
  }
  parseParaProperties(r) {
    const e = r.getElementsByTagName("hh:paraPr");
    for (let n = 0; n < e.length; n++) {
      const i = e[n], a = i.getAttribute("id") || "", s = (Z) => Array.from(i.childNodes).find(
        (T) => T.nodeType === 1 && T.localName === Z
      );
      let o = 0, h = 0, l = 0, f = 0, d = 0;
      const u = s("margin");
      if (u) {
        o = parseInt(u.getAttribute("left") || "0", 10), h = parseInt(u.getAttribute("right") || "0", 10), f = parseInt(u.getAttribute("prev") || "0", 10), d = parseInt(u.getAttribute("next") || "0", 10);
        const Z = Array.from(u.childNodes).find(
          (T) => T.nodeType === 1 && T.localName === "indent"
        );
        Z && (l = parseInt(Z.getAttribute("value") || "0", 10));
      }
      let c = "PERCENT", g = 160;
      const p = s("lineSpacing");
      p && (c = p.getAttribute("type") || "PERCENT", g = parseInt(p.getAttribute("value") || "160", 10));
      const y = s("align"), m = (y == null ? void 0 : y.getAttribute("value")) || "JUSTIFY", A = i.getAttribute("keepWithNext") === "true" ? "1" : "0", E = i.getAttribute("keepLines") === "true" ? "1" : "0", x = i.getAttribute("pageBreakBefore") === "true" ? "1" : "0", S = s("heading"), R = (S == null ? void 0 : S.getAttribute("type")) || "NONE", v = (S == null ? void 0 : S.getAttribute("level")) || "0", O = (S == null ? void 0 : S.getAttribute("idRef")) || "0", C = s("borderFill"), B = (C == null ? void 0 : C.getAttribute("borderFillIDRef")) || "";
      this.ctx.paraProperties.set(a, {
        id: a,
        align: m,
        heading: R,
        level: v,
        headingIdRef: O,
        leftMargin: o,
        rightMargin: h,
        indent: l,
        prevSpacing: f,
        nextSpacing: d,
        lineSpacingType: c,
        lineSpacingVal: g,
        keepWithNext: A,
        keepLines: E,
        pageBreakBefore: x,
        borderFillIDRef: B
      });
    }
  }
  parseStyles(r) {
  }
  parseNumberings(r) {
  }
}
let Oi = class {
  constructor(r) {
    ce(this, "ctx");
    this.ctx = r;
  }
  convertSection(r) {
    const i = new DOMParser().parseFromString(r, "text/xml").documentElement;
    let a = "";
    this.extractPageSetup(i);
    const s = i.childNodes;
    for (let o = 0; o < s.length; o++) {
      const h = s[o];
      h.nodeType === 1 && (h.localName === "p" ? a += this.convertParagraph(h) : h.localName === "tbl" && (a += this.convertTable(h)));
    }
    return a;
  }
  extractPageSetup(r) {
    const e = r.getElementsByTagNameNS(this.ctx.ns.hp, "pagePr");
    if (!e || e.length === 0) return;
    const n = e[0], i = parseInt(n.getAttribute("width") || "0", 10), a = parseInt(n.getAttribute("height") || "0", 10);
    i > 0 && (this.ctx.pageWidth = Math.round(i / 5)), a > 0 && (this.ctx.pageHeight = Math.round(a / 5));
    const s = Array.from(n.childNodes).find((g) => g.nodeType === 1 && g.localName === "margin");
    if (!s) return;
    const o = (g) => parseInt(s.getAttribute(g) || "0", 10), h = o("top"), l = o("bottom"), f = o("left"), d = o("right"), u = o("header"), c = o("footer");
    h > 0 && (this.ctx.marginTop = Math.round(h / 5)), l > 0 && (this.ctx.marginBottom = Math.round(l / 5)), f > 0 && (this.ctx.marginLeft = Math.round(f / 5)), d > 0 && (this.ctx.marginRight = Math.round(d / 5)), this.ctx.marginHeader = u > 0 ? Math.round(u / 5) : this.ctx.marginTop, this.ctx.marginFooter = c > 0 ? Math.round(c / 5) : this.ctx.marginBottom;
  }
  convertParagraph(r) {
    const e = r.getAttribute("paraPrIDRef") || "0", n = r.getAttribute("pageBreak") || "0";
    let i = "<w:p>";
    i += this.getParaPropertiesXml(e, n);
    const a = r.getElementsByTagNameNS(this.ctx.ns.hp, "run");
    for (const s of Array.from(a)) {
      if (s.parentNode !== r) continue;
      const o = s.getAttribute("charPrIDRef") || "0", l = Array.from(s.getElementsByTagNameNS(this.ctx.ns.hp, "tbl")).filter((g) => {
        let p = g.parentNode;
        for (; p && p !== s; ) {
          if (p.localName === "run") return !1;
          p = p.parentNode;
        }
        return p === s;
      });
      if (l.length > 0) {
        i += "</w:p>";
        for (const g of l)
          i += this.convertTable(g);
        i += "<w:p>", i += this.getParaPropertiesXml(e, "0");
        continue;
      }
      const f = s.getElementsByTagNameNS(this.ctx.ns.hp, "pic");
      if (f.length > 0) {
        for (const g of Array.from(f))
          i += this.convertImage(g, o);
        continue;
      }
      const d = s.getElementsByTagNameNS(this.ctx.ns.hp, "rect");
      if (d.length > 0) {
        for (const g of Array.from(d))
          i += this.convertRect(g, o);
        continue;
      }
      if (s.getElementsByTagNameNS(this.ctx.ns.hp, "secPr").length > 0) continue;
      const c = s.getElementsByTagNameNS(this.ctx.ns.hp, "t");
      if (c.length > 0)
        for (const g of Array.from(c)) {
          const p = this.getTextContent(g);
          p && (i += "<w:r>", i += this.getRunPropertiesXml(o), i += `<w:t xml:space="preserve">${this.escapeXml(p)}</w:t>`, i += "</w:r>");
          const y = g.getElementsByTagNameNS(this.ctx.ns.hp, "tab");
          for (const m of Array.from(y))
            i += "<w:r>", i += this.getRunPropertiesXml(o), i += "<w:tab/>", i += "</w:r>";
        }
    }
    return i += "</w:p>", i;
  }
  getTextContent(r) {
    let e = "";
    for (const n of Array.from(r.childNodes))
      n.nodeType === 3 && (e += n.nodeValue);
    return e;
  }
  getParaPropertiesXml(r, e) {
    const n = this.ctx.paraProperties.get(r);
    if (!n && e !== "1") return "";
    let i = "<w:pPr>";
    if (n) {
      const a = { LEFT: "left", CENTER: "center", RIGHT: "right", JUSTIFY: "both", DISTRIBUTE: "distribute" };
      if (n.align && n.align !== "LEFT" && (i += `<w:jc w:val="${a[n.align] || "left"}"/>`), n.leftMargin || n.rightMargin || n.indent) {
        let o = "<w:ind";
        n.leftMargin && (o += ` w:left="${Math.round(n.leftMargin / 5)}"`), n.rightMargin && (o += ` w:right="${Math.round(n.rightMargin / 5)}"`), n.indent > 0 ? o += ` w:firstLine="${Math.round(n.indent / 5)}"` : n.indent < 0 && (o += ` w:hanging="${Math.round(Math.abs(n.indent) / 5)}"`), o += "/>", i += o;
      }
      let s = "<w:spacing";
      n.prevSpacing && (s += ` w:before="${Math.round(n.prevSpacing / 5)}"`), n.nextSpacing ? s += ` w:after="${Math.round(n.nextSpacing / 5)}"` : s += ' w:after="0"', n.lineSpacingType === "PERCENT" ? (s += ` w:line="${Math.round(n.lineSpacingVal * 240 / 100)}"`, s += ' w:lineRule="auto"') : n.lineSpacingType === "FIXED" ? (s += ` w:line="${Math.round(n.lineSpacingVal / 5)}"`, s += ' w:lineRule="exact"') : (n.lineSpacingType === "BETWEENLINES" || n.lineSpacingType === "AT_LEAST") && (s += ` w:line="${Math.round(n.lineSpacingVal / 5)}"`, s += ' w:lineRule="atLeast"'), s += "/>", i += s, n.keepWithNext === "1" && (i += "<w:keepNext/>"), n.keepLines === "1" && (i += "<w:keepLines/>");
    }
    return e === "1" && (i += "<w:pageBreakBefore/>"), i += "</w:pPr>", i;
  }
  getRunPropertiesXml(r) {
    const e = this.ctx.charProperties.get(r);
    if (!e) return "";
    let n = "<w:rPr>";
    (e.latinFont || e.hangulFont) && (n += `<w:rFonts w:ascii="${e.latinFont}" w:eastAsia="${e.hangulFont}" w:hAnsi="${e.latinFont}"/>`), e.bold && (n += "<w:b/>"), e.italic && (n += "<w:i/>"), e.underline && (n += '<w:u w:val="single"/>'), e.strikeout && (n += "<w:strike/>"), e.supscript && (n += '<w:vertAlign w:val="superscript"/>'), e.subscript && (n += '<w:vertAlign w:val="subscript"/>');
    const i = Math.round(e.height / 50);
    if (i > 0 && (n += `<w:sz w:val="${i}"/>`, n += `<w:szCs w:val="${i}"/>`), e.textColor && e.textColor !== "#000000" && (n += `<w:color w:val="${e.textColor.replace("#", "")}"/>`), e.shadeColor && e.shadeColor !== "none")
      n += `<w:highlight w:val="${this.colorToHighlightName(e.shadeColor)}"/>`;
    else if (e.borderFillIDRef && e.borderFillIDRef !== "1") {
      const a = this.ctx.borderFills.get(e.borderFillIDRef);
      a && a.faceColor && a.faceColor !== "#FFFFFF" && a.faceColor !== "none" && (n += `<w:shd w:val="clear" w:color="auto" w:fill="${a.faceColor.replace("#", "")}"/>`);
    }
    return n += "</w:rPr>", n;
  }
  colorToHighlightName(r) {
    return {
      "#FFFF00": "yellow",
      "#00FFFF": "cyan",
      "#00FF00": "green",
      "#FF00FF": "magenta",
      "#0000FF": "blue",
      "#FF0000": "red",
      "#000000": "black",
      "#FFFFFF": "white",
      "#008000": "darkGreen",
      "#808000": "darkYellow",
      "#800000": "darkRed",
      "#008080": "darkCyan",
      "#800080": "darkMagenta",
      "#00008B": "darkBlue",
      "#A9A9A9": "darkGray",
      "#D3D3D3": "lightGray"
    }[r.toUpperCase()] || "yellow";
  }
  convertTable(r) {
    let e = "<w:tbl>";
    const n = r.getElementsByTagNameNS(this.ctx.ns.hp, "sz")[0], i = n ? parseInt(n.getAttribute("width") || "0", 10) : 0, a = i > 0 ? Math.round(i / 5) : 0;
    let s = "left";
    const o = r.getElementsByTagNameNS(this.ctx.ns.hp, "pos")[0];
    if (o && o.getAttribute("treatAsChar") !== "1") {
      const A = o.getAttribute("horzAlign") || "LEFT";
      A === "CENTER" ? s = "center" : A === "RIGHT" ? s = "right" : s = "left";
    }
    const h = r.getAttribute("borderFillIDRef") || "", l = this.ctx.borderFills.get(h);
    e += "<w:tblPr>", e += a > 0 ? `<w:tblW w:w="${a}" w:type="dxa"/>` : '<w:tblW w:w="0" w:type="auto"/>', e += `<w:jc w:val="${s}"/>`, e += this.buildTblBordersXml(l), e += '<w:tblLayout w:type="fixed"/>', e += "</w:tblPr>";
    const f = Array.from(r.childNodes).filter((A) => A.nodeType === 1 && A.localName === "tr"), d = [];
    for (const A of f) {
      const E = Array.from(A.childNodes).filter((x) => x.nodeType === 1 && x.localName === "tc");
      for (const x of E) {
        const S = Array.from(x.childNodes).find((T) => T.nodeType === 1 && T.localName === "cellAddr"), R = Array.from(x.childNodes).find((T) => T.nodeType === 1 && T.localName === "cellSpan"), v = Array.from(x.childNodes).find((T) => T.nodeType === 1 && T.localName === "cellSz") || Array.from(x.childNodes).find((T) => T.nodeType === 1 && T.localName === "sz"), O = S ? parseInt(S.getAttribute("colAddr") || "0", 10) : 0, C = R ? Math.max(1, parseInt(R.getAttribute("colSpan") || "1", 10)) : 1, B = v ? parseInt(v.getAttribute("width") || "0", 10) : 0, Z = Math.round(B / 5);
        d.push({ colAddr: O, colSpan: C, widthDxa: Z });
      }
    }
    let u = 0;
    for (const A of d)
      u = Math.max(u, A.colAddr + A.colSpan);
    const c = /* @__PURE__ */ new Map();
    for (const A of d)
      A.colSpan === 1 && A.widthDxa > 0 && !c.has(A.colAddr) && c.set(A.colAddr, A.widthDxa);
    for (const A of d)
      if (A.colSpan > 1 && A.widthDxa > 0) {
        const E = [];
        let x = 0;
        for (let S = A.colAddr; S < A.colAddr + A.colSpan; S++)
          c.has(S) ? x += c.get(S) : E.push(S);
        if (E.length > 0) {
          const S = Math.max(0, A.widthDxa - x), R = Math.round(S / E.length);
          for (const v of E)
            c.set(v, R);
        }
      }
    const g = [];
    if (u > 0)
      for (let A = 0; A < u; A++)
        g.push(c.get(A) || 0);
    if (g.reduce((A, E) => A + E, 0) === 0 && a > 0) {
      const A = g.length || 1, E = Math.round(a / A);
      g.length === 0 ? g.push(a) : g.fill(E);
    } else g.length === 0 && a > 0 && g.push(a);
    e += "<w:tblGrid>";
    for (const A of g) e += `<w:gridCol w:w="${A}"/>`;
    e += "</w:tblGrid>";
    const y = /* @__PURE__ */ new Set(), m = g.length;
    for (let A = 0; A < f.length; A++) {
      const E = f[A], x = Array.from(E.childNodes).filter((v) => v.nodeType === 1 && v.localName === "tc");
      if (e += "<w:tr>", x.length > 0) {
        const v = Array.from(x[0].childNodes).find((C) => C.nodeType === 1 && C.localName === "cellSz"), O = v ? parseInt(v.getAttribute("height") || "0", 10) : 0;
        O > 0 && (e += `<w:trPr><w:trHeight w:val="${Math.round(O / 5)}"/></w:trPr>`);
      }
      let S = 0, R = 0;
      for (; S < m || R < x.length; ) {
        if (S >= m) {
          R++;
          continue;
        }
        if (y.has(`${A},${S}`))
          e += "<w:tc>", e += "<w:tcPr>", e += `<w:tcW w:w="${g[S] || 0}" w:type="dxa"/>`, e += "<w:vMerge/>", e += "</w:tcPr>", e += "<w:p/>", e += "</w:tc>", S++;
        else if (R < x.length) {
          const v = x[R++], O = Array.from(v.childNodes).find((T) => T.nodeType === 1 && T.localName === "cellSpan"), C = O ? Math.max(1, parseInt(O.getAttribute("colSpan") || "1", 10)) : 1, B = O ? Math.max(1, parseInt(O.getAttribute("rowSpan") || "1", 10)) : 1;
          if (B > 1)
            for (let T = 1; T < B; T++)
              for (let j = 0; j < C; j++)
                y.add(`${A + T},${S + j}`);
          let Z = 0;
          for (let T = 0; T < C; T++)
            Z += g[S + T] || 0;
          e += this.convertTableCell(v, C, B, Z), S += C;
        } else
          e += `<w:tc><w:tcPr><w:tcW w:w="${g[S] || 0}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`, S++;
      }
      e += "</w:tr>";
    }
    return e += "</w:tbl>", e;
  }
  convertTableCell(r, e = 1, n = 1, i = 0) {
    let a = "<w:tc>";
    a += "<w:tcPr>";
    let s = i;
    if (s === 0) {
      const p = Array.from(r.childNodes).find((m) => m.nodeType === 1 && m.localName === "cellSz") || Array.from(r.childNodes).find((m) => m.nodeType === 1 && m.localName === "sz"), y = p ? parseInt(p.getAttribute("width") || "0", 10) : 0;
      s = Math.round(y / 5);
    }
    a += `<w:tcW w:w="${s}" w:type="dxa"/>`, e > 1 && (a += `<w:gridSpan w:val="${e}"/>`), n > 1 && (a += '<w:vMerge w:val="restart"/>');
    const o = r.getAttribute("borderFillIDRef") || "", h = this.ctx.borderFills.get(o);
    if (a += this.buildTcBordersXml(h), h != null && h.faceColor) {
      const p = h.faceColor.replace("#", "");
      a += `<w:shd w:val="clear" w:color="auto" w:fill="${p}"/>`;
    }
    const l = Array.from(r.childNodes).find((p) => p.nodeType === 1 && p.localName === "cellMargin");
    if (l) {
      const p = (x) => Math.round(parseInt(l.getAttribute(x) || "0", 10) / 5), y = p("left"), m = p("right"), A = p("top"), E = p("bottom");
      a += "<w:tcMar>", a += `<w:top w:w="${A}" w:type="dxa"/>`, a += `<w:left w:w="${y}" w:type="dxa"/>`, a += `<w:bottom w:w="${E}" w:type="dxa"/>`, a += `<w:right w:w="${m}" w:type="dxa"/>`, a += "</w:tcMar>";
    }
    const f = Array.from(r.childNodes).find((p) => p.nodeType === 1 && p.localName === "subList"), d = (f == null ? void 0 : f.getAttribute("vertAlign")) || "CENTER";
    a += `<w:vAlign w:val="${{ TOP: "top", CENTER: "center", BOTTOM: "bottom" }[d] || "center"}"/>`, a += "</w:tcPr>";
    const c = r.getElementsByTagNameNS(this.ctx.ns.hp, "subList")[0];
    let g = !1;
    if (c) {
      const p = Array.from(c.childNodes).filter((y) => y.nodeType === 1 && y.localName === "p");
      for (const y of p)
        a += this.convertParagraph(y), g = !0;
    }
    return g || (a += "<w:p/>"), a += "</w:tc>", a;
  }
  convertImage(r, e) {
    const n = r.getElementsByTagNameNS(this.ctx.ns.hc, "img")[0];
    if (!n) return "";
    const i = n.getAttribute("binaryItemIDRef");
    if (!i || !this.ctx.images.has(i)) return "";
    const a = this.ctx.images.get(i);
    if (!a) return "";
    const s = r.getElementsByTagNameNS(this.ctx.ns.hp, "sz")[0], o = (this.ctx.pageWidth - this.ctx.marginLeft - this.ctx.marginRight) * 635;
    let h = Math.round(o / 2), l = Math.round(h * 3 / 4);
    if (s) {
      const g = parseInt(s.getAttribute("width") || "0"), p = parseInt(s.getAttribute("height") || "0");
      g > 0 && (h = g * 127), p > 0 && (l = p * 127);
    }
    const f = `rId${++this.ctx.relIdCounter}`, d = this.ctx.relIdCounter, u = `image${d}.${a.ext}`;
    this.ctx.imageRels.push({
      id: f,
      target: `media/${u}`,
      data: a.data,
      mediaType: a.mediaType,
      ext: a.ext
    });
    let c = "<w:r>";
    return c += this.getRunPropertiesXml(e), c += "<w:drawing>", c += '<wp:inline distT="0" distB="0" distL="0" distR="0">', c += `<wp:extent cx="${h}" cy="${l}"/>`, c += `<wp:docPr id="${d}" name="Picture ${d}"/>`, c += '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">', c += '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">', c += '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">', c += `<pic:nvPicPr><pic:cNvPr id="${d}" name="Picture ${d}"/><pic:cNvPicPr/></pic:nvPicPr>`, c += `<pic:blipFill><a:blip r:embed="${f}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`, c += `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${h}" cy="${l}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`, c += "</pic:pic>", c += "</a:graphicData>", c += "</a:graphic>", c += "</wp:inline>", c += "</w:drawing>", c += "</w:r>", c;
  }
  convertRect(r, e) {
    let n = "";
    const i = r.getElementsByTagNameNS(this.ctx.ns.hp, "drawText")[0];
    if (!i) return "";
    const a = i.getElementsByTagNameNS(this.ctx.ns.hp, "subList")[0];
    if (!a) return "";
    const s = a.getElementsByTagNameNS(this.ctx.ns.hp, "p");
    for (const o of Array.from(s)) {
      if (o.parentNode !== a) continue;
      const h = o.getElementsByTagNameNS(this.ctx.ns.hp, "run");
      for (const l of Array.from(h)) {
        if (l.parentNode !== o) continue;
        const f = l.getAttribute("charPrIDRef") || e, d = l.getElementsByTagNameNS(this.ctx.ns.hp, "t");
        for (const u of Array.from(d)) {
          const c = this.getTextContent(u);
          c && (n += "<w:r>", n += this.getRunPropertiesXml(f), n += `<w:t xml:space="preserve">${this.escapeXml(c)}</w:t>`, n += "</w:r>");
        }
      }
    }
    return n;
  }
  // ─── 표 테두리 보조 함수 ───────────────────────────────────────
  /**
   * HWPX border type → DOCX w:val 변환
   * HWPX width(mm 문자열) → DOCX sz (1/8 포인트 단위)
   */
  hwpBorderToDocx(r, e, n) {
    const a = {
      NONE: "none",
      SOLID: "single",
      DOUBLE: "double",
      DASHED: "dashed",
      DOTTED: "dotted",
      DOUBLE_THIN: "double"
    }[r] || "none", s = parseFloat(e), o = isNaN(s) ? 2 : Math.max(2, Math.round(s * 72 / 25.4 * 8)), h = n ? n.replace("#", "") : "000000";
    return { val: a, sz: o, color: h };
  }
  /** tbl 레벨 <w:tblBorders> 생성 */
  buildTblBordersXml(r) {
    const e = ["top", "left", "bottom", "right"], n = {
      top: r == null ? void 0 : r.topBorder,
      left: r == null ? void 0 : r.leftBorder,
      bottom: r == null ? void 0 : r.bottomBorder,
      right: r == null ? void 0 : r.rightBorder
    };
    let i = "<w:tblBorders>";
    for (const s of e) {
      const o = n[s];
      if (o && o.type !== "NONE") {
        const { val: h, sz: l, color: f } = this.hwpBorderToDocx(o.type, o.width, o.color);
        i += `<w:${s} w:val="${h}" w:sz="${l}" w:space="0" w:color="${f}"/>`;
      } else
        i += `<w:${s} w:val="nil"/>`;
    }
    const a = r ? [r.topBorder, r.leftBorder, r.bottomBorder, r.rightBorder].filter((s) => s && s.type !== "NONE") : [];
    if (a.length > 0) {
      const s = a[0], { val: o, sz: h, color: l } = this.hwpBorderToDocx(s.type, s.width, s.color);
      i += `<w:insideH w:val="${o}" w:sz="${h}" w:space="0" w:color="${l}"/>`, i += `<w:insideV w:val="${o}" w:sz="${h}" w:space="0" w:color="${l}"/>`;
    } else
      i += '<w:insideH w:val="nil"/>', i += '<w:insideV w:val="nil"/>';
    return i += "</w:tblBorders>", i;
  }
  /** tc 레벨 <w:tcBorders> 생성 */
  buildTcBordersXml(r) {
    if (!r) return "";
    const e = [
      { key: "topBorder", name: "top" },
      { key: "leftBorder", name: "left" },
      { key: "bottomBorder", name: "bottom" },
      { key: "rightBorder", name: "right" }
    ];
    let n = "<w:tcBorders>";
    for (const { key: i, name: a } of e) {
      const s = r[i];
      if (s) {
        const { val: o, sz: h, color: l } = this.hwpBorderToDocx(s.type, s.width, s.color);
        n += `<w:${a} w:val="${o}" w:sz="${h}" w:space="0" w:color="${l}"/>`;
      }
    }
    return n += "</w:tcBorders>", n;
  }
  escapeXml(r) {
    return r ? r.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;") : "";
  }
};
const Ir = class Ir {
  constructor(r) {
    ce(this, "ctx");
    this.ctx = r;
  }
  async createDocxPackage(r) {
    const e = new yt();
    e.file("[Content_Types].xml", this.createContentTypes()), e.file("_rels/.rels", this.createTopRels()), e.file("word/document.xml", this.createDocumentXml(r)), e.file("word/_rels/document.xml.rels", this.createDocumentRels()), e.file("word/styles.xml", this.createStylesXml()), e.file("word/settings.xml", this.createSettingsXml()), e.file("word/fontTable.xml", this.createFontTableXml());
    for (const n of this.ctx.imageRels)
      e.file(`word/${n.target}`, n.data);
    return await e.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  }
  // ─── 한국어 폰트 판별 ────────────────────────────────────
  isKoreanFont(r) {
    return /함초롬|맑은|나눔|굴림|바탕|돋움|궁서|HCR|새굴림|윤|한컴|[가-힣]/.test(r);
  }
  // ─── 문서에서 사용된 폰트 목록 수집 ────────────────────────
  getDocumentFonts() {
    const r = /* @__PURE__ */ new Set();
    for (const e of this.ctx.charProperties.values())
      e.hangulFont && r.add(e.hangulFont), e.latinFont && r.add(e.latinFont);
    for (const e of Object.values(this.ctx.fontFaces))
      for (const n of e)
        n.face && r.add(n.face);
    return [...r].filter((e) => e.trim());
  }
  // ─── 기본 폰트 결정 ─────────────────────────────────────
  getDefaultFonts() {
    const r = this.getDocumentFonts(), e = r.find((i) => this.isKoreanFont(i)) ?? "Malgun Gothic", n = r.find((i) => !this.isKoreanFont(i) && /^[A-Za-z]/.test(i)) ?? e;
    return { korean: e, latin: n };
  }
  createContentTypes() {
    const r = /* @__PURE__ */ new Set();
    for (const n of this.ctx.imageRels)
      n.ext && r.add(n.ext.toLowerCase());
    let e = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    e += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">', e += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>', e += '<Default Extension="xml" ContentType="application/xml"/>';
    for (const n of r) {
      const i = Ir.MIME_MAP[n] ?? `image/${n}`;
      e += `<Default Extension="${n}" ContentType="${i}"/>`;
    }
    return e += '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>', e += '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>', e += '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>', e += '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>', e += "</Types>", e;
  }
  createTopRels() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
  }
  createDocumentXml(r) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${r}<w:sectPr><w:pgSz w:w="${this.ctx.pageWidth}" w:h="${this.ctx.pageHeight}"/><w:pgMar w:top="${this.ctx.marginTop}" w:right="${this.ctx.marginRight}" w:bottom="${this.ctx.marginBottom}" w:left="${this.ctx.marginLeft}" w:header="${this.ctx.marginHeader}" w:footer="${this.ctx.marginFooter}" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  }
  createDocumentRels() {
    let r = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    r += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">', r += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>', r += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>', r += '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>';
    for (const e of this.ctx.imageRels)
      r += `<Relationship Id="${e.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${e.target.replace("word/", "")}"/>`;
    return r += "</Relationships>", r;
  }
  /** styles.xml: 문서의 실제 폰트를 기본값으로 사용 */
  createStylesXml() {
    const { korean: r, latin: e } = this.getDefaultFonts();
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${e}" w:eastAsia="${r}" w:hAnsi="${e}" w:cs="${r}"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults></w:styles>`;
  }
  createSettingsXml() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/></w:settings>';
  }
  /** fontTable.xml: 문서에 사용된 폰트 전체 목록 등록 */
  createFontTableXml() {
    const r = this.getDocumentFonts();
    r.length === 0 && r.push("Malgun Gothic");
    let e = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    e += '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';
    for (const n of r) {
      const i = this.isKoreanFont(n) ? "81" : "00";
      e += `<w:font w:name="${n}"><w:charset w:val="${i}"/><w:family w:val="auto"/><w:pitch w:val="variable"/></w:font>`;
    }
    return e += "</w:fonts>", e;
  }
};
// ─── MIME 타입 매핑 ──────────────────────────────────────
ce(Ir, "MIME_MAP", {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  bmp: "image/bmp",
  gif: "image/gif",
  tiff: "image/tiff",
  tif: "image/tiff",
  emf: "image/x-emf",
  wmf: "image/x-wmf",
  svg: "image/svg+xml",
  webp: "image/webp"
});
let Qr = Ir;
class Hi {
  async convert(r) {
    let e;
    return r instanceof File || r instanceof Blob ? e = await yt.loadAsync(r) : e = await yt.loadAsync(r), await this.convertFromZip(e);
  }
  async convertFromZip(r) {
    console.log("HWPX ZIP에서 변환...");
    const e = new Li(), n = new Bi(e), i = new Oi(e), a = new Qr(e), s = r.file("Contents/content.hpf");
    s && await n.parseContentHpf(await s.async("string"));
    const o = r.file("Contents/header.xml");
    o && n.parseHeader(await o.async("string"));
    const h = r.folder("BinData");
    if (h) {
      const d = [];
      h.forEach((u, c) => {
        c.dir || d.push({ name: u, file: c });
      });
      for (const { name: u, file: c } of d) {
        const g = await c.async("uint8array"), p = u.split("."), y = p.length > 1 ? p.pop().toLowerCase() : "", m = p.length > 0 ? p[0].split("/").pop() : u, A = y === "png" ? "image/png" : y === "jpg" || y === "jpeg" ? "image/jpeg" : `image/${y}`;
        e.images.set(m, { data: g, mediaType: A, ext: y });
      }
    }
    const l = Object.keys(r.files).filter((d) => /Contents\/section\d+\.xml$/i.test(d)).sort((d, u) => {
      var p, y;
      const c = parseInt(((p = d.match(/section(\d+)/i)) == null ? void 0 : p[1]) || "0", 10), g = parseInt(((y = u.match(/section(\d+)/i)) == null ? void 0 : y[1]) || "0", 10);
      return c - g;
    });
    let f = "";
    for (const d of l) {
      const u = r.file(d);
      u && (f += i.convertSection(await u.async("string")));
    }
    return await a.createDocxPackage(f);
  }
}
const Di = async (t) => await new Hi().convert(t), lh = () => {
  const [t, r] = Oe(!1), [e, n] = Oe(null), [i, a] = Oe(null);
  return { convert: async (o) => {
    r(!0), n(null), a(null);
    try {
      const h = await ki(o);
      return a(h), h;
    } catch (h) {
      throw n(h), h;
    } finally {
      r(!1);
    }
  }, isLoading: t, error: e, result: i };
}, hh = () => {
  const [t, r] = Oe(!1), [e, n] = Oe(null), [i, a] = Oe(null);
  return { convert: async (o) => {
    r(!0), n(null), a(null);
    try {
      const h = await Di(o);
      return a(h), h;
    } catch (h) {
      throw n(h), h;
    } finally {
      r(!1);
    }
  }, isLoading: t, error: e, result: i };
};
function $t(t) {
  return t ? t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;") : "";
}
function Mi(t, r = "", e = 0) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="no"?>
<HWPML Style="embed" SubVersion="10.0.0.0" Version="2.91">
    <HEAD SecCnt="1">
        <DOCSUMMARY>
            <TITLE>Report_</TITLE>
            <AUTHOR>___AUTHOR___</AUTHOR>
            <DATE>___DATE___</DATE>
        </DOCSUMMARY>
        <DOCSETTING>
            <BEGINNUMBER Endnote="1" Equation="1" Footnote="1" Page="1" Picture="1" Table="1" />
            <CARETPOS List="0" Para="36" Pos="0" />
        </DOCSETTING>
        <MAPPINGTABLE>
            ${t}
            <FACENAMELIST>
                <FONTFACE Count="2" Lang="Hangul">
                    <FONT Id="0" Name="함초롬돋움" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                    <FONT Id="1" Name="함초롬바탕" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                </FONTFACE>
                <FONTFACE Count="2" Lang="Latin">
                    <FONT Id="0" Name="함초롬돋움" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                    <FONT Id="1" Name="함초롬바탕" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                </FONTFACE>
                <FONTFACE Count="2" Lang="Hanja">
                    <FONT Id="0" Name="함초롬돋움" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                    <FONT Id="1" Name="함초롬바탕" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                </FONTFACE>
                <FONTFACE Count="2" Lang="Japanese">
                    <FONT Id="0" Name="함초롬돋움" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                    <FONT Id="1" Name="함초롬바탕" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                </FONTFACE>
                <FONTFACE Count="2" Lang="Other">
                    <FONT Id="0" Name="함초롬돋움" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                    <FONT Id="1" Name="함초롬바탕" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                </FONTFACE>
                <FONTFACE Count="2" Lang="Symbol">
                    <FONT Id="0" Name="함초롬돋움" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                    <FONT Id="1" Name="함초롬바탕" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                </FONTFACE>
                <FONTFACE Count="2" Lang="User">
                    <FONT Id="0" Name="함초롬돋움" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                    <FONT Id="1" Name="함초롬바탕" Type="ttf">
                        <TYPEINFO ArmStyle="0" Contrast="0" FamilyType="0" Letterform="0"
                            Midline="252" Proportion="0" StrokeVariation="0" Weight="0"
                            XHeight="255" />
                    </FONT>
                </FONTFACE>
            </FACENAMELIST>
            <BORDERFILLLIST Count="${4 + e}">
                <BORDERFILL BackSlash="0" BreakCellSeparateLine="0" CenterLine="0"
                    CounterBackSlash="0" CounterSlash="0" CrookedSlash="0" Id="1" Shadow="false"
                    Slash="0" ThreeD="false">
                    <LEFTBORDER Type="None" Width="0.1mm" />
                    <RIGHTBORDER Type="None" Width="0.1mm" />
                    <TOPBORDER Type="None" Width="0.1mm" />
                    <BOTTOMBORDER Type="None" Width="0.1mm" />
                    <DIAGONAL Type="Solid" Width="0.1mm" />
                </BORDERFILL>
                <BORDERFILL BackSlash="0" BreakCellSeparateLine="0" CenterLine="0"
                    CounterBackSlash="0" CounterSlash="0" CrookedSlash="0" Id="2" Shadow="false"
                    Slash="0" ThreeD="false">
                    <LEFTBORDER Type="None" Width="0.1mm" />
                    <RIGHTBORDER Type="None" Width="0.1mm" />
                    <TOPBORDER Type="None" Width="0.1mm" />
                    <BOTTOMBORDER Type="None" Width="0.1mm" />
                    <DIAGONAL Type="Solid" Width="0.1mm" />
                    <FILLBRUSH>
                        <WINDOWBRUSH Alpha="0" FaceColor="4294967295" HatchColor="10066329" />
                    </FILLBRUSH>
                </BORDERFILL>
                <BORDERFILL BackSlash="0" BreakCellSeparateLine="0" CenterLine="0"
                    CounterBackSlash="0" CounterSlash="0" CrookedSlash="0" Id="3" Shadow="false"
                    Slash="0" ThreeD="false">
                    <LEFTBORDER Type="Solid" Width="0.12mm" />
                    <RIGHTBORDER Type="Solid" Width="0.12mm" />
                    <TOPBORDER Type="Solid" Width="0.12mm" />
                    <BOTTOMBORDER Type="Solid" Width="0.12mm" />
                    <DIAGONAL Type="Solid" Width="0.1mm" />
                </BORDERFILL>
                <BORDERFILL BackSlash="0" BreakCellSeparateLine="0" CenterLine="0"
                    CounterBackSlash="0" CounterSlash="0" CrookedSlash="0" Id="4" Shadow="false"
                    Slash="0" ThreeD="false">
                    <LEFTBORDER Type="Solid" Width="0.12mm" />
                    <RIGHTBORDER Type="Solid" Width="0.12mm" />
                    <TOPBORDER Type="Solid" Width="0.12mm" />
                    <BOTTOMBORDER Type="Solid" Width="0.12mm" />
                    <DIAGONAL Type="Solid" Width="0.1mm" />
                    <FILLBRUSH>
                        <WINDOWBRUSH Alpha="0" FaceColor="15132390" HatchColor="4294967295" />
                    </FILLBRUSH>
                </BORDERFILL>
                ${r}
            </BORDERFILLLIST>
            <CHARSHAPELIST Count="20">
                <CHARSHAPE BorderFillId="2" Height="1000" Id="0" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1000" Id="1" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="900" Id="2" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="900" Id="3" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="900" Id="4" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="-5" Hanja="-5" Japanese="-5" Latin="-5" Other="-5"
                        Symbol="-5" User="-5" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1600" Id="5" ShadeColor="4294967295" SymMark="0"
                    TextColor="11891758" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1100" Id="6" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1000" Id="7" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <BOLD />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1900" Id="8" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="2400" Id="9" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="3200" Id="10" ShadeColor="4294967295"
                    SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <BOLD />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="2400" Id="11" ShadeColor="4294967295"
                    SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <BOLD />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1900" Id="12" ShadeColor="4294967295"
                    SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <BOLD />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1600" Id="13" ShadeColor="4294967295"
                    SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <BOLD />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1200" Id="14" ShadeColor="4294967295"
                    SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <BOLD />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1000" Id="15" ShadeColor="4294967295"
                    SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <ITALIC />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1000" Id="16" ShadeColor="4294967295"
                    SymMark="0" TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <BOLD />
                    <UNDERLINE Color="0" Shape="Solid" Type="Bottom" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1000" Id="17" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <STRIKEOUT Color="0" Shape="Solid" Type="Continuous" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1000" Id="18" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <UNDERLINE Color="0" Shape="Solid" Type="Bottom" />
                </CHARSHAPE>
                <CHARSHAPE BorderFillId="2" Height="1000" Id="19" ShadeColor="4294967295" SymMark="0"
                    TextColor="0" UseFontSpace="false" UseKerning="false">
                    <FONTID Hangul="1" Hanja="1" Japanese="1" Latin="1" Other="1" Symbol="1"
                        User="1" />
                    <RATIO Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHARSPACING Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <RELSIZE Hangul="100" Hanja="100" Japanese="100" Latin="100" Other="100"
                        Symbol="100" User="100" />
                    <CHAROFFSET Hangul="0" Hanja="0" Japanese="0" Latin="0" Other="0" Symbol="0"
                        User="0" />
                    <BOLD />
                    <ITALIC />
                </CHARSHAPE>
            </CHARSHAPELIST>
            <TABDEFLIST Count="3">
                <TABDEF AutoTabLeft="false" AutoTabRight="false" Id="0" />
                <TABDEF AutoTabLeft="true" AutoTabRight="false" Id="1" />
                <TABDEF AutoTabLeft="false" AutoTabRight="true" Id="2" />
            </TABDEFLIST>
            <NUMBERINGLIST Count="2">
                <NUMBERING Id="1" Start="0">
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="1" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="2"
                        NumFormat="HangulSyllable" Start="1" TextOffset="50"
                        TextOffsetType="percent" UseInstWidth="true" WidthAdjust="0">^2.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="3" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^3)</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="4"
                        NumFormat="HangulSyllable" Start="1" TextOffset="50"
                        TextOffsetType="percent" UseInstWidth="true" WidthAdjust="0">^4)</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="5" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">(^5)</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="6"
                        NumFormat="HangulSyllable" Start="1" TextOffset="50"
                        TextOffsetType="percent" UseInstWidth="true" WidthAdjust="0">(^6)</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="7" NumFormat="CircledDigit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^7</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="8"
                        NumFormat="CircledHangulSyllable" Start="1" TextOffset="50"
                        TextOffsetType="percent" UseInstWidth="true" WidthAdjust="0">^8</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="9" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0"></PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="10" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0"></PARAHEAD>
                </NUMBERING>
                <NUMBERING Id="2" Start="0">
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="1" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="2" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.^2.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="3" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.^2.^3.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="4" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.^2.^3.^4.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="5" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.^2.^3.^4.^5.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="6" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.^2.^3.^4.^5.^6.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="7" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.^2.^3.^4.^5.^6.^7.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="8" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.^2.^3.^4.^5.^6.^7.^8.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="9" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.^2.^3.^4.^5.^6.^7.^8.^9.</PARAHEAD>
                    <PARAHEAD Alignment="Left" AutoIndent="true" Level="10" NumFormat="Digit"
                        Start="1" TextOffset="50" TextOffsetType="percent" UseInstWidth="true"
                        WidthAdjust="0">^1.^2.^3.^4.^5.^6.^7.^8.^9.^:.</PARAHEAD>
                </NUMBERING>
            </NUMBERINGLIST>
            <BULLETLIST Count="1">
                <BULLET Char="" Id="1">
                    <PARAHEAD Alignment="Left" AutoIndent="true" TextOffset="50"
                        TextOffsetType="percent" UseInstWidth="false" WidthAdjust="0" />
                </BULLET>
            </BULLETLIST>
            <PARASHAPELIST Count="27">
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="0" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="1" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="3000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="20"
                    FontLineHeight="false" HeadingType="Outline" Id="2" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="2000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="20"
                    FontLineHeight="false" HeadingType="Outline" Id="3" KeepLines="false"
                    KeepWithNext="false" Level="1" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="4000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="20"
                    FontLineHeight="false" HeadingType="Outline" Id="4" KeepLines="false"
                    KeepWithNext="false" Level="2" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="6000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="20"
                    FontLineHeight="false" HeadingType="Outline" Id="5" KeepLines="false"
                    KeepWithNext="false" Level="3" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="8000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="20"
                    FontLineHeight="false" HeadingType="Outline" Id="6" KeepLines="false"
                    KeepWithNext="false" Level="4" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="10000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="20"
                    FontLineHeight="false" HeadingType="Outline" Id="7" KeepLines="false"
                    KeepWithNext="false" Level="5" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="12000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="20"
                    FontLineHeight="false" HeadingType="Outline" Id="8" KeepLines="false"
                    KeepWithNext="false" Level="6" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="14000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="false" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="9" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="150" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="10" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="-2620" Left="0" LineSpacing="130" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Left" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="false" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="11" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="130" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Left" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="false" Condense="20"
                    FontLineHeight="false" HeadingType="None" Id="12" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="160" LineSpacingType="Percent"
                        Next="600" Prev="2400" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Left" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="false" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="13" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="2" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="160" LineSpacingType="Percent"
                        Next="1400" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Left" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="false" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="14" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="2" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="2200" LineSpacing="160" LineSpacingType="Percent"
                        Next="1400" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Left" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="false" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="15" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="2" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="4400" LineSpacing="160" LineSpacingType="Percent"
                        Next="1400" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" HeadingType="Outline" Id="16" KeepLines="false"
                    KeepWithNext="false" Level="8" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="18000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" HeadingType="Outline" Id="17" KeepLines="false"
                    KeepWithNext="false" Level="9" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="20000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" HeadingType="Outline" Id="18" KeepLines="false"
                    KeepWithNext="false" Level="7" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="16000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="19" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="150" LineSpacingType="Percent"
                        Next="1600" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="20" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" Heading="1" HeadingType="Bullet" Id="21"
                    KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break"
                    PageBreakBefore="false" SnapToGrid="true" TabDef="0" VerAlign="Baseline"
                    WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="20"
                    FontLineHeight="false" Heading="1" HeadingType="Bullet" Id="22"
                    KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break"
                    PageBreakBefore="false" SnapToGrid="true" TabDef="1" VerAlign="Baseline"
                    WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="2000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="23" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="20"
                    FontLineHeight="false" HeadingType="None" Id="24" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="1" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="2000" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Justify" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" Heading="1" HeadingType="Number" Id="25"
                    KeepLines="false" KeepWithNext="false" Level="0" LineWrap="Break"
                    PageBreakBefore="false" SnapToGrid="true" TabDef="0" VerAlign="Baseline"
                    WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
                <PARASHAPE Align="Center" AutoSpaceEAsianEng="false" AutoSpaceEAsianNum="false"
                    BreakLatinWord="KeepWord" BreakNonLatinWord="true" Condense="0"
                    FontLineHeight="false" HeadingType="None" Id="26" KeepLines="false"
                    KeepWithNext="false" Level="0" LineWrap="Break" PageBreakBefore="false"
                    SnapToGrid="true" TabDef="0" VerAlign="Baseline" WidowOrphan="false">
                    <PARAMARGIN Indent="0" Left="0" LineSpacing="160" LineSpacingType="Percent"
                        Next="0" Prev="0" Right="0" />
                    <PARABORDER BorderFill="2" Connect="false" IgnoreMargin="false" />
                </PARASHAPE>
            </PARASHAPELIST>
            <STYLELIST Count="22">
                <STYLE CharShape="0" EngName="Normal" Id="0" LangId="1042" LockForm="0" Name="바탕글"
                    NextStyle="0" ParaShape="0" Type="Para" />
                <STYLE CharShape="0" EngName="Body" Id="1" LangId="1042" LockForm="0" Name="본문"
                    NextStyle="1" ParaShape="1" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 1" Id="2" LangId="1042" LockForm="0"
                    Name="개요 1" NextStyle="2" ParaShape="2" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 2" Id="3" LangId="1042" LockForm="0"
                    Name="개요 2" NextStyle="3" ParaShape="3" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 3" Id="4" LangId="1042" LockForm="0"
                    Name="개요 3" NextStyle="4" ParaShape="4" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 4" Id="5" LangId="1042" LockForm="0"
                    Name="개요 4" NextStyle="5" ParaShape="5" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 5" Id="6" LangId="1042" LockForm="0"
                    Name="개요 5" NextStyle="6" ParaShape="6" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 6" Id="7" LangId="1042" LockForm="0"
                    Name="개요 6" NextStyle="7" ParaShape="7" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 7" Id="8" LangId="1042" LockForm="0"
                    Name="개요 7" NextStyle="8" ParaShape="8" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 8" Id="9" LangId="1042" LockForm="0"
                    Name="개요 8" NextStyle="9" ParaShape="18" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 9" Id="10" LangId="1042" LockForm="0"
                    Name="개요 9" NextStyle="10" ParaShape="16" Type="Para" />
                <STYLE CharShape="0" EngName="Outline 10" Id="11" LangId="1042" LockForm="0"
                    Name="개요 10" NextStyle="11" ParaShape="17" Type="Para" />
                <STYLE CharShape="1" EngName="Page Number" Id="12" LangId="1042" LockForm="0"
                    Name="쪽 번호" NextStyle="0" Type="Char" />
                <STYLE CharShape="2" EngName="Header" Id="13" LangId="1042" LockForm="0" Name="머리말"
                    NextStyle="13" ParaShape="9" Type="Para" />
                <STYLE CharShape="3" EngName="Footnote" Id="14" LangId="1042" LockForm="0" Name="각주"
                    NextStyle="14" ParaShape="10" Type="Para" />
                <STYLE CharShape="3" EngName="Endnote" Id="15" LangId="1042" LockForm="0" Name="미주"
                    NextStyle="15" ParaShape="10" Type="Para" />
                <STYLE CharShape="4" EngName="Memo" Id="16" LangId="1042" LockForm="0" Name="메모"
                    NextStyle="16" ParaShape="11" Type="Para" />
                <STYLE CharShape="5" EngName="TOC Heading" Id="17" LangId="1042" LockForm="0"
                    Name="차례 제목" NextStyle="17" ParaShape="12" Type="Para" />
                <STYLE CharShape="6" EngName="TOC 1" Id="18" LangId="1042" LockForm="0" Name="차례 1"
                    NextStyle="18" ParaShape="13" Type="Para" />
                <STYLE CharShape="6" EngName="TOC 2" Id="19" LangId="1042" LockForm="0" Name="차례 2"
                    NextStyle="19" ParaShape="14" Type="Para" />
                <STYLE CharShape="6" EngName="TOC 3" Id="20" LangId="1042" LockForm="0" Name="차례 3"
                    NextStyle="20" ParaShape="15" Type="Para" />
                <STYLE CharShape="0" EngName="Caption" Id="21" LangId="1042" LockForm="0" Name="캡션"
                    NextStyle="21" ParaShape="19" Type="Para" />
            </STYLELIST>
        </MAPPINGTABLE>
        <COMPATIBLEDOCUMENT TargetProgram="None">
            <LAYOUTCOMPATIBILITY AdjustBaselineInFixedLinespacing="false"
                AdjustBaselineOfObjectToBottom="false" AdjustLineheightToFont="false"
                AdjustMarginFromAdjustLineheight="false" AdjustParaBorderOffsetWithBorder="false"
                AdjustParaBorderfillToSpacing="false" AdjustVertPosOfLine="false"
                ApplyAtLeastToPercent100Pct="false" ApplyCharSpacingToCharGrid="false"
                ApplyExtendHeaderFooterEachSection="false" ApplyFontWeightToBold="false"
                ApplyFontspaceToLatin="false" ApplyMinColumnWidthTo1mm="false"
                ApplyNextspacingOfLastPara="false" ApplyParaBorderToOutside="false"
                ApplyPrevspacingBeneathObject="false" ApplyTabPosBasedOnSegment="false"
                BaseCharUnitOfIndentOnFirstChar="false" BaseCharUnitOnEAsian="false"
                BaseLinespacingOnLinegrid="false" BreakTabOverLine="false"
                ConnectParaBorderfillOfEqualBorder="false" DoNotAdjustEmptyAnchorLine="false"
                DoNotAdjustWordInJustify="false" DoNotAlignLastForbidden="false"
                DoNotAlignLastPeriod="false" DoNotAlignWhitespaceOnRight="false"
                DoNotApplyAutoSpaceEAsianEng="false" DoNotApplyAutoSpaceEAsianNum="false"
                DoNotApplyColSeparatorAtNoGap="false" DoNotApplyExtensionCharCompose="false"
                DoNotApplyGridInHeaderFooter="false" DoNotApplyHeaderFooterAtNoSpace="false"
                DoNotApplyImageEffect="false" DoNotApplyLinegridAtNoLinespacing="false"
                DoNotApplyShapeComment="false" DoNotApplyStrikeoutWithUnderline="false"
                DoNotApplyVertOffsetOfForward="false" DoNotApplyWhiteSpaceHeight="false"
                DoNotFormattingAtBeneathAnchor="false" DoNotHoldAnchorOfTable="false"
                ExtendLineheightToOffset="false" ExtendLineheightToParaBorderOffset="false"
                ExtendVertLimitToPageMargins="false" FixedUnderlineWidth="false"
                OverlapBothAllowOverlap="false" TreatQuotationAsLatin="false"
                UseInnerUnderline="false" UseLowercaseStrikeout="false" />
        </COMPATIBLEDOCUMENT>
    </HEAD>
    <BODY>
        <SECTION Id="0">
            <SECDEF CharGrid="0" FirstBorder="false" FirstFill="false" LineGrid="0" SpaceBetweenColumns="0" TabStop="8000" TextDirection="0" VerticalPageAlign="Top">
                <STARTNUMBER Endnote="1" Equation="1" Footnote="1" Page="1" Picture="1" Table="1"/>
                <PAGEDEF GutterType="Left" Height="84188" Landscape="0" Width="59528">
                    <PAGEMARGIN Bottom="4252" Footer="4252" Gutter="0" Header="4252" Left="8504" Right="8504" Top="5668"/>
                </PAGEDEF>
            </SECDEF>`;
}
const $i = '<P ParaShape="20" Style="0"><TEXT CharShape="10"><CHAR>', Ui = '<P ParaShape="20" Style="0"><TEXT CharShape="11"><CHAR>', zi = '<P ParaShape="20" Style="0"><TEXT CharShape="12"><CHAR>', Wi = '<P ParaShape="20" Style="0"><TEXT CharShape="13"><CHAR>', ji = '<P ParaShape="20" Style="0"><TEXT CharShape="14"><CHAR>', Gi = '<P ParaShape="20" Style="0"><TEXT CharShape="7"><CHAR>', Xi = `
`;
class Zi {
  constructor(r) {
    ce(this, "fs");
    ce(this, "BIN_ITEM_ENTRIES", []);
    ce(this, "BIN_STORAGE_ENTRIES", []);
    ce(this, "PLACEHOLDERS", {});
    ce(this, "ph_counter", 0);
    ce(this, "borderFills", []);
    this.fs = r;
  }
  register_placeholder(r) {
    const e = `__MD2HML_PH_${this.ph_counter}__`;
    return this.PLACEHOLDERS[e] = r, this.ph_counter++, e;
  }
  async fetchImage(r) {
    try {
      const e = await fetch(r);
      if (!e.ok) throw new Error(`HTTP ${e.status}`);
      const n = await e.arrayBuffer(), i = new Uint8Array(n);
      let a = "jpg";
      const s = e.headers.get("content-type");
      if (s)
        s.includes("png") ? a = "png" : s.includes("gif") ? a = "gif" : s.includes("bmp") && (a = "bmp");
      else {
        const o = r.split("."), h = o.length > 1 ? o[o.length - 1].toLowerCase() : "";
        h && (a = h);
      }
      return { data: i, ext: a };
    } catch (e) {
      throw e;
    }
  }
  async getImageSize(r) {
    return new Promise((e) => {
      try {
        if (typeof window > "u") {
          e({ width: 0, height: 0 });
          return;
        }
        const n = new Blob([r]), i = URL.createObjectURL(n), a = new Image();
        a.onload = () => {
          const s = a.naturalWidth, o = a.naturalHeight;
          URL.revokeObjectURL(i), e({ width: s, height: o });
        }, a.onerror = () => {
          URL.revokeObjectURL(i), e({ width: 0, height: 0 });
        }, a.src = i;
      } catch (n) {
        console.warn("Image size calculation failed:", n), e({ width: 0, height: 0 });
      }
    });
  }
  async process_image(r, e) {
    let n = new Uint8Array(0), i = "jpg";
    try {
      if (e.startsWith("data:")) {
        const c = e.indexOf(",");
        if (c > -1) {
          const g = e.substring(c + 1), p = e.substring(5, c).split(";")[0];
          if (typeof Buffer < "u")
            n = new Uint8Array(Buffer.from(g, "base64"));
          else {
            const y = atob(g), m = y.length;
            n = new Uint8Array(m);
            for (let A = 0; A < m; A++) n[A] = y.charCodeAt(A);
          }
          p.includes("png") ? i = "png" : p.includes("gif") ? i = "gif" : p.includes("bmp") ? i = "bmp" : i = "jpg";
        }
      } else if (e.startsWith("http://") || e.startsWith("https://")) {
        const c = await this.fetchImage(e);
        if (!c) throw new Error("Fetch failed");
        n = c.data, i = c.ext;
      } else {
        let c = !1;
        if (await this.fs.exists(e))
          n = await this.fs.readFile(e), c = !0;
        else {
          const p = e.split("/").pop() || e;
          await this.fs.exists(p) && (n = await this.fs.readFile(p), c = !0);
        }
        if (!c)
          return `<P ParaShape="0" Style="0"><TEXT CharShape="0"><CHAR>[Image not found: ${e}]</CHAR></TEXT></P>`;
        const g = e.split(".");
        i = g.length > 1 ? g[g.length - 1].toLowerCase() : "jpg";
      }
      i === "jpeg" && (i = "jpg"), ["jpg", "png", "gif", "bmp"].includes(i) || (i = "jpg");
      let s = "";
      if (typeof Buffer < "u")
        s = Buffer.from(n).toString("base64");
      else {
        let c = "";
        const g = n.byteLength;
        for (let p = 0; p < g; p++)
          c += String.fromCharCode(n[p]);
        s = window.btoa(c);
      }
      const o = s.length, h = this.BIN_ITEM_ENTRIES.length + 1;
      this.BIN_ITEM_ENTRIES.push(
        `<BINITEM BinData="${h}" Format="${i}" Type="Embedding" />`
      ), this.BIN_STORAGE_ENTRIES.push(
        `<BINDATA Encoding="Base64" Id="${h}" Size="${o}">${s}</BINDATA>`
      );
      let l = 28346, f = 28346;
      try {
        const c = await this.getImageSize(n);
        if (c.width && c.height) {
          l = Math.floor(c.width * 75), f = Math.floor(c.height * 75);
          const g = 42e3;
          if (l > g) {
            const p = g / l;
            l = g, f = Math.floor(f * p);
          }
        }
      } catch (c) {
        console.warn("Image size error (using default):", c);
      }
      const d = 1e7 + h;
      return `<P ParaShape="0" Style="0"><TEXT CharShape="0"><PICTURE Reverse="false">
<SHAPEOBJECT Lock="false" NumberingType="Figure" TextFlow="BothSides" TextWrap="Square" ZOrder="6">
<SIZE Height="${f}" HeightRelTo="Absolute" Protect="false" Width="${l}" WidthRelTo="Absolute"/>
<POSITION AffectLSpacing="false" AllowOverlap="true" FlowWithText="true" HoldAnchorAndSO="false" HorzAlign="Left" HorzOffset="0" HorzRelTo="Column" TreatAsChar="true" VertAlign="Top" VertOffset="0" VertRelTo="Para"/>
<OUTSIDEMARGIN Bottom="0" Left="0" Right="0" Top="0"/>
<SHAPECOMMENT>${$t(r)}</SHAPECOMMENT>
</SHAPEOBJECT>
<SHAPECOMPONENT GroupLevel="0" HorzFlip="false" InstID="${d}" OriHeight="${f}" OriWidth="${l}" VertFlip="false" XPos="0" YPos="0">
<ROTATIONINFO Angle="0" CenterX="${Math.floor(l / 2)}" CenterY="${Math.floor(f / 2)}" Rotate="1"/>
<RENDERINGINFO>
<TRANSMATRIX E1="1.00000" E2="0.00000" E3="0.00000" E4="0.00000" E5="1.00000" E6="0.00000"/>
<SCAMATRIX E1="1.00000" E2="0.00000" E3="0.00000" E4="0.00000" E5="1.00000" E6="0.00000"/>
<ROTMATRIX E1="1.00000" E2="0.00000" E3="0.00000" E4="0.00000" E5="1.00000" E6="0.00000"/>
</RENDERINGINFO>
</SHAPECOMPONENT>
<IMAGERECT X0="0" X1="${l}" X2="${l}" X3="0" Y0="0" Y1="0" Y2="${f}" Y3="${f}"/>
<IMAGECLIP Bottom="${f}" Left="0" Right="${l}" Top="0"/>
<INSIDEMARGIN Bottom="0" Left="0" Right="0" Top="0"/>
<IMAGE Alpha="0" BinItem="${h}" Bright="0" Contrast="0" Effect="RealPic"/>
<EFFECTS/>
</PICTURE></TEXT></P>`;
    } catch (a) {
      return `<P ParaShape="0" Style="0"><TEXT CharShape="0"><CHAR>[Image error: ${a.message}]</CHAR></TEXT></P>`;
    }
  }
  process_code_block(r) {
    const e = r.trim().split(`
`);
    let n = "";
    return e.forEach((i) => {
      n += `<P ParaShape="0" Style="0"><TEXT CharShape="16"><CHAR>${$t(i)}</CHAR></TEXT></P>`;
    }), `<P ParaShape="0" Style="0"><TEXT CharShape="0"><TABLE BorderFill="3" CellSpacing="0" ColCount="1" PageBreak="Cell" RepeatHeader="1" RowCount="1">
<SHAPEOBJECT Lock="0" NumberingType="Table" TextWrap="TopAndBottom" ZOrder="1">
<SIZE Height="0" HeightRelTo="Absolute" Protect="0" Width="41954" WidthRelTo="Absolute"/>
<POSITION AffectLSpacing="0" AllowOverlap="0" FlowWithText="1" HoldAnchorAndSO="0" HorzAlign="Left" HorzOffset="0" HorzRelTo="Column" TreatAsChar="0" VertAlign="Top" VertOffset="0" VertRelTo="Para"/>
<OUTSIDEMARGIN Bottom="283" Left="283" Right="283" Top="283"/>
</SHAPEOBJECT>
<INSIDEMARGIN Bottom="141" Left="510" Right="510" Top="141"/>
<ROW>
<CELL BorderFill="3" ColAddr="0" ColSpan="1" Dirty="0" Editable="0" HasMargin="0" Header="0" Height="282" Protect="0" RowAddr="0" RowSpan="1" Width="41954">
<CELLMARGIN Bottom="141" Left="510" Right="510" Top="141"/>
<PARALIST LineWrap="Break" LinkListID="0" LinkListIDNext="0" TextDirection="0" VertAlign="Center">
${n}
</PARALIST>
</CELL>
</ROW>
</TABLE><CHAR/></TEXT></P>`;
  }
  process_table_native(r) {
    const e = r.trim().split(`
`);
    if (e.length < 2) return r;
    const n = e[0].trim().replace(/^\||\|$/g, "").split("|").map((f) => f.trim()), i = n.length, a = [];
    a.push(n);
    for (let f = 2; f < e.length; f++) {
      let d = e[f].trim().replace(/^\||\|$/g, "").split("|").map((u) => u.trim());
      d.length < i && (d = d.concat(
        new Array(i - d.length).fill("")
      )), a.push(d.slice(0, i));
    }
    const s = a.length, o = 41954, h = Math.floor(o / i);
    let l = "";
    return a.forEach((f, d) => {
      let u = "";
      f.forEach((c, g) => {
        const p = d === 0 ? 4 : 3, y = d === 0 ? 1 : 0, m = d === 0 ? 26 : 0, A = d === 0 ? 7 : 0, E = c.split(/<br\s*\/?>|\n/i);
        let x = "";
        E.forEach((S) => {
          x += `<P ParaShape="${m}" Style="0"><TEXT CharShape="${A}"><CHAR>${$t(S.trim() || " ")}</CHAR></TEXT></P>
`;
        }), u += `<CELL BorderFill="${p}" ColAddr="${g}" ColSpan="1" Dirty="0" Editable="0" HasMargin="0" Header="${y}" Height="282" Protect="0" RowAddr="${d}" RowSpan="1" Width="${h}">
<CELLMARGIN Bottom="141" Left="510" Right="510" Top="141"/>
<PARALIST LineWrap="Break" LinkListID="0" LinkListIDNext="0" TextDirection="0" VertAlign="Center">
${x.trim()}
</PARALIST>
</CELL>`;
      }), l += `<ROW>${u}</ROW>`;
    }), `<P ParaShape="0" Style="0"><TEXT CharShape="0"><TABLE BorderFill="3" CellSpacing="0" ColCount="${i}" PageBreak="Cell" RepeatHeader="1" RowCount="${s}">
<SHAPEOBJECT Lock="0" NumberingType="Table" TextWrap="TopAndBottom" ZOrder="0">
<SIZE Height="0" HeightRelTo="Absolute" Protect="0" Width="${o}" WidthRelTo="Absolute"/>
<POSITION AffectLSpacing="0" AllowOverlap="0" FlowWithText="1" HoldAnchorAndSO="0" HorzAlign="Left" HorzOffset="0" HorzRelTo="Column" TreatAsChar="0" VertAlign="Top" VertOffset="0" VertRelTo="Para"/>
<OUTSIDEMARGIN Bottom="283" Left="283" Right="283" Top="283"/>
</SHAPEOBJECT>
<INSIDEMARGIN Bottom="141" Left="510" Right="510" Top="141"/>
${l}
</TABLE><CHAR/></TEXT></P>`;
  }
  processInline(r) {
    return r = r.replace(
      new RegExp("\\*\\*\\*(?=\\S)([\\s\\S]+?)(?<=\\S)\\*\\*\\*", "g"),
      (e, n) => `</CHAR></TEXT><TEXT CharShape="19"><CHAR>${n}</CHAR></TEXT><TEXT CharShape="0"><CHAR>`
    ), r = r.replace(new RegExp("___(?=\\S)([\\s\\S]+?)(?<=\\S)___", "g"), (e, n) => `</CHAR></TEXT><TEXT CharShape="19"><CHAR>${n}</CHAR></TEXT><TEXT CharShape="0"><CHAR>`), r = r.replace(new RegExp("\\*\\*(?=\\S)([\\s\\S]+?)(?<=\\S)\\*\\*", "g"), (e, n) => `</CHAR></TEXT><TEXT CharShape="7"><CHAR>${n}</CHAR></TEXT><TEXT CharShape="0"><CHAR>`), r = r.replace(
      new RegExp("(?<!\\*)\\*(?=\\S)([\\s\\S]+?)(?<=\\S)\\*(?!\\*)", "g"),
      (e, n) => `</CHAR></TEXT><TEXT CharShape="15"><CHAR>${n}</CHAR></TEXT><TEXT CharShape="0"><CHAR>`
    ), r = r.replace(new RegExp("__(?=\\S)([\\s\\S]+?)(?<=\\S)__", "g"), (e, n) => `</CHAR></TEXT><TEXT CharShape="7"><CHAR>${n}</CHAR></TEXT><TEXT CharShape="0"><CHAR>`), r = r.replace(
      new RegExp("(?<!_)_(?=\\S)([\\s\\S]+?)(?<=\\S)_(?!_)", "g"),
      (e, n) => `</CHAR></TEXT><TEXT CharShape="15"><CHAR>${n}</CHAR></TEXT><TEXT CharShape="0"><CHAR>`
    ), r = r.replace(new RegExp("~~(?=\\S)([\\s\\S]+?)(?<=\\S)~~", "g"), (e, n) => `</CHAR></TEXT><TEXT CharShape="17"><CHAR>${n}</CHAR></TEXT><TEXT CharShape="0"><CHAR>`), r = r.replace(/&lt;u&gt;([\s\S]*?)&lt;\/u&gt;/gi, (e, n) => `</CHAR></TEXT><TEXT CharShape="18"><CHAR>${n}</CHAR></TEXT><TEXT CharShape="0"><CHAR>`), r = r.replace(/`([^`]+)`/g, (e, n) => `</CHAR></TEXT><TEXT CharShape="16"><CHAR>${n}</CHAR></TEXT><TEXT CharShape="0"><CHAR>`), r;
  }
  async convert(r) {
    let e = r;
    e = e.replace(/\r/g, ""), e = `
` + e, this.BIN_ITEM_ENTRIES = [], this.BIN_STORAGE_ENTRIES = [], this.PLACEHOLDERS = {}, this.ph_counter = 0, this.borderFills = [];
    const n = /!\[(.*?)\]\((.*?)\)/g;
    let i;
    const a = [];
    for (; (i = n.exec(e)) !== null; )
      a.push({
        fullMatch: i[0],
        alt: i[1],
        path: i[2]
      });
    const s = await Promise.all(
      a.map(async (S) => {
        const R = await this.process_image(S.alt, S.path), v = this.register_placeholder(R);
        return { fullMatch: S.fullMatch, placeholder: v };
      })
    );
    let o = 0;
    e = e.replace(n, () => {
      const S = s[o];
      return o++, S ? S.placeholder : "";
    }), e = e.replace(
      /^```([\s\S]*?)```/gm,
      (S, R) => this.register_placeholder(this.process_code_block(R))
    );
    let h = "";
    for (; e !== h; )
      h = e, e = e.replace(
        /<table(?:(?!<table)[\s\S])*?<\/table>/gi,
        (S) => this.register_placeholder(this.process_html_table(S))
      );
    e = $t(e);
    let l = "No Title", f = "Unknown", d = "Unknown", u = e.match(new RegExp("(?<=title: ).*"));
    u && (l = u[0]);
    let c = e.match(new RegExp("(?<=author: ).*"));
    c && (f = c[0]);
    let g = e.match(new RegExp("(?<=date: ).*"));
    g && (d = g[0]), e = e.replace(/---[\s\S]*?---/, ""), e = e.replace(
      /\n---/g,
      `
<P ParaShape="10" Style="0"><TEXT CharShape="0"><CHAR></CHAR></TEXT></P>`
    ), e = e.replace(/\[(.*?)\]\((.*?)\)/g, "$1 ($2)");
    let p = 3102933706;
    e = e.replace(
      /\n( {0,24})([-+*]) /g,
      (S, R, v) => {
        const C = 21 + Math.floor(R.length / 4);
        let B = "";
        return C === 22 && (p++, B = ` InstId="${p}"`), `
<P${B} ParaShape="${C}" Style="0"><TEXT CharShape="0"><CHAR>`;
      }
    ), e = e.replace(
      /\n(\d+)\. /g,
      `
<P ParaShape="25" Style="0"><TEXT CharShape="0"><CHAR>`
    ), e = e.replace(
      /\n    (\d+)\. /g,
      `
<P InstId="95545006$1" ParaShape="2" Style="2"><TEXT CharShape="0"><CHAR>`
    ), e = e.replace(/\[ \]/g, "☐"), e = e.replace(/\[x\]/gi, "☑"), e = e.replace(
      /\n(&gt;|>) (.*)/g,
      `
<P ParaShape="11" Style="0"><TEXT CharShape="8"><CHAR>$2</CHAR></TEXT></P>`
    ), e = e.replace(
      /((\n?\|.*\|\s*)+)/gm,
      (S, R) => this.process_table_native(R)
    );
    for (let S = this.ph_counter - 1; S >= 0; S--) {
      const R = `__MD2HML_PH_${S}__`;
      this.PLACEHOLDERS.hasOwnProperty(R) && (e = e.split(R).join(this.PLACEHOLDERS[R]));
    }
    e = e.replace(/\n# /g, `
` + $i), e = e.replace(/\n## /g, `
` + Ui), e = e.replace(/\n### /g, `
` + zi), e = e.replace(/\n#### /g, `
` + Wi), e = e.replace(/\n##### /g, `
` + ji), e = e.replace(/\n###### /g, `
` + Gi), e = e.replace(
      /\n(?=[^<])/g,
      `
<P ParaShape="0" Style="0"><TEXT CharShape="0"><CHAR>`
    ), e = e.replace(
      /(<CHAR>)([^<]*?)(?=\s*(?:<\/TEXT>|<\/P>|<P|$))/gs,
      "$1$2</CHAR>"
    ), e = e.replace(/(<\/CHAR>)(?!\s*<\/TEXT>)/g, "$1</TEXT></P>"), e = e.replace(/(<CHAR>[^<]*)$/gm, "$1</CHAR></TEXT></P>"), e = e.replace(
      /(<TEXT CharShape="0"><CHAR>)(.*?)(<\/CHAR>)/gs,
      (S, R, v, O) => R + this.processInline(v) + O
    ), e = e.replace(
      /(<TEXT CharShape="0"><CHAR>)(.*?)(<\/CHAR>)/gs,
      (S, R, v, O) => R + this.processInline(v) + O
    );
    let y = "";
    this.BIN_ITEM_ENTRIES.length > 0 && (y = `<BINDATALIST Count="${this.BIN_ITEM_ENTRIES.length}">` + this.BIN_ITEM_ENTRIES.join("") + "</BINDATALIST>");
    const m = this.borderFills.map((S) => S.xml).join(`
`), A = this.HEADER(
      y,
      m,
      this.borderFills.length
    );
    let E = "";
    this.BIN_STORAGE_ENTRIES.length > 0 && (E = "<BINDATASTORAGE>" + this.BIN_STORAGE_ENTRIES.join("") + "</BINDATASTORAGE>");
    const x = `</SECTION></BODY><TAIL>${E}</TAIL></HWPML>`;
    return e = A + Xi + e + x, e = e.replace("___TITLE___", l), e = e.replace("___AUTHOR___", f), e = e.replace("___DATE___", d), e;
  }
  HEADER(r, e = "", n = 0) {
    return Mi(
      r,
      e,
      n
    );
  }
  process_html_table(r) {
    const e = [];
    let n = 0;
    const i = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
    let a;
    for (; (a = i.exec(r)) !== null; ) {
      const g = a[1], p = /<(td|th)([^>]*)>([\s\S]*?)<\/\1>/gi;
      let y;
      const m = [];
      for (; (y = p.exec(g)) !== null; ) {
        const A = y[1].toLowerCase(), E = y[2], x = y[3].trim();
        let S = 1, R = 1, v = A;
        const O = E.match(/colspan=["'](\d+)["']/i);
        O && (S = parseInt(O[1]) || 1);
        const C = E.match(/rowspan=["'](\d+)["']/i);
        C && (R = parseInt(C[1]) || 1);
        const B = E.match(/style=["']([^"']+)["']/i);
        B && (v += " " + B[1]), m.push({ text: x, style: v, colspan: S, rowspan: R });
      }
      if (m.length > n) {
        let A = 0;
        m.forEach((E) => A += E.colspan), A > n && (n = A);
      }
      m.length > 0 && e.push(m);
    }
    if (e.length === 0) return r;
    const s = [];
    let o = 0;
    e.forEach((g, p) => {
      s[p] || (s[p] = []);
      let y = 0;
      g.forEach((m) => {
        for (; s[p][y]; )
          y++;
        for (let A = 0; A < m.rowspan; A++) {
          const E = p + A;
          s[E] || (s[E] = []);
          for (let x = 0; x < m.colspan; x++)
            s[E][y + x] = !0;
        }
        o = Math.max(o, y + m.colspan), m.colAddr = y, y += m.colspan;
      });
    }), o > n && (n = o);
    const h = 41954, l = new Array(n).fill(0);
    let f = !1;
    if (e.length > 0) {
      let g = 0;
      for (const p of e[0]) {
        const m = p.style.replace(/^(th|td)\s*/, "").match(/(?:^|;)\s*width\s*:\s*(\d+)\s*px/i);
        if (m) {
          f = !0;
          const A = parseInt(m[1]) / p.colspan;
          for (let E = 0; E < p.colspan && g + E < n; E++)
            l[g + E] = A;
        }
        g += p.colspan;
      }
    }
    let d;
    if (f && l.every((g) => g > 0)) {
      const g = l.reduce((y, m) => y + m, 0);
      d = l.map(
        (y) => Math.floor(y / g * h)
      );
      const p = d.reduce((y, m) => y + m, 0);
      d[n - 1] += h - p;
    } else {
      const g = Math.floor(h / n);
      d = new Array(n).fill(g), d[n - 1] += h - g * n;
    }
    let u = "";
    e.forEach((g, p) => {
      let y = "";
      g.forEach((m) => {
        const A = m.style.startsWith("th") || p === 0;
        let E = A ? 4 : 3;
        const x = A ? 1 : 0, S = A ? 26 : 0, R = A ? 7 : 0, v = m.colAddr, O = m.style.replace(/^(th|td)\s*/, "");
        if (O.trim()) {
          const T = this.parseCssToHmlBorderFill(O);
          if (T) {
            const j = 4 + this.borderFills.length + 1;
            this.borderFills.push({ id: j, xml: T }), E = j;
          }
        }
        let C = "";
        m.text.match(/^__MD2HML_PH_\d+__$/) ? C = `${m.text}
` : m.text.split(/<br\s*\/?>|\n/i).forEach((j) => {
          C += `<P ParaShape="${S}" Style="0"><TEXT CharShape="${R}"><CHAR>${$t(j.trim() || " ")}</CHAR></TEXT></P>
`;
        });
        const B = d.slice(v, v + m.colspan).reduce((T, j) => T + j, 0);
        y += `<CELL BorderFill="${E}" ColAddr="${v}" ColSpan="${m.colspan}" Dirty="0" Editable="0" HasMargin="0" Header="${x}" Height="282" Protect="0" RowAddr="${p}" RowSpan="${m.rowspan}" Width="${B}">
<CELLMARGIN Bottom="141" Left="510" Right="510" Top="141"/>
<PARALIST LineWrap="Break" LinkListID="0" LinkListIDNext="0" TextDirection="0" VertAlign="Center">
${C.trim()}
</PARALIST>
</CELL>`;
      }), u += `<ROW>${y}</ROW>`;
    });
    const c = e.length;
    return `<P ParaShape="0" Style="0"><TEXT CharShape="0"><TABLE BorderFill="3" CellSpacing="0" ColCount="${n}" PageBreak="Cell" RepeatHeader="1" RowCount="${c}">
<SHAPEOBJECT Lock="0" NumberingType="Table" TextWrap="TopAndBottom" ZOrder="0">
<SIZE Height="0" HeightRelTo="Absolute" Protect="0" Width="${h}" WidthRelTo="Absolute"/>
<POSITION AffectLSpacing="0" AllowOverlap="0" FlowWithText="1" HoldAnchorAndSO="0" HorzAlign="Left" HorzOffset="0" HorzRelTo="Column" TreatAsChar="0" VertAlign="Top" VertOffset="0" VertRelTo="Para"/>
<OUTSIDEMARGIN Bottom="283" Left="283" Right="283" Top="283"/>
</SHAPEOBJECT>
<INSIDEMARGIN Bottom="141" Left="510" Right="510" Top="141"/>
${u}
</TABLE><CHAR/></TEXT></P>`;
  }
  /**
   * CSS style 문자열에서 border/background 정보를 파싱하여
   * HML BORDERFILL XML을 생성합니다.
   */
  parseCssToHmlBorderFill(r) {
    const e = (u) => {
      if (new RegExp(
        `${u}\\s*:\\s*(?:none|0(?:px)?)(?:\\s*;|\\s*$|\\s)`,
        "i"
      ).test(r))
        return { width: "0.1mm", color: "#000000", none: !0 };
      const g = new RegExp(
        `${u}\\s*:\\s*(\\d+(?:\\.\\d+)?)px\\s+\\w+\\s+([#\\w]+)`,
        "i"
      ), p = r.match(g);
      return p ? { width: (parseFloat(p[1]) * 25.4 / 96).toFixed(2) + "mm", color: p[2] } : null;
    }, n = e("border-left") || e("border"), i = e("border-right") || e("border"), a = e("border-top") || e("border"), s = e("border-bottom") || e("border");
    let o = null;
    const h = r.match(
      /(?:background|background-color)\s*:\s*([#\w]+)/i
    );
    if (h) {
      const u = h[1];
      if (u.startsWith("#") && u.length >= 7) {
        const c = parseInt(u.slice(1, 3), 16), g = parseInt(u.slice(3, 5), 16), p = parseInt(u.slice(5, 7), 16);
        o = c | g << 8 | p << 16;
      }
    }
    if (!n && !i && !a && !s && o === null) return null;
    const l = (u, c) => !c || c.none ? `<${u} Type="None" Width="0.1mm" />` : `<${u} Color="${this.cssColorToHmlInt(c.color)}" Type="Solid" Width="${c.width}" />`;
    let d = `<BORDERFILL BackSlash="0" BreakCellSeparateLine="0" CenterLine="0" CounterBackSlash="0" CounterSlash="0" CrookedSlash="0" Id="${4 + this.borderFills.length + 1}" Shadow="false" Slash="0" ThreeD="false">
`;
    return d += l("LEFTBORDER", n) + `
`, d += l("RIGHTBORDER", i) + `
`, d += l("TOPBORDER", a) + `
`, d += l("BOTTOMBORDER", s) + `
`, d += `<DIAGONAL Type="Solid" Width="0.1mm" />
`, o !== null && (d += `<FILLBRUSH>
<WINDOWBRUSH Alpha="0" FaceColor="${o}" HatchColor="4294967295" />
</FILLBRUSH>
`), d += "</BORDERFILL>", d;
  }
  cssColorToHmlInt(r) {
    if (r.startsWith("#") && r.length >= 7) {
      const e = parseInt(r.slice(1, 3), 16), n = parseInt(r.slice(3, 5), 16), i = parseInt(r.slice(5, 7), 16);
      return e | n << 8 | i << 16;
    }
    return 0;
  }
}
var wa = { exports: {} };
const Ki = {}, Ji = /* @__PURE__ */ Object.freeze(/* @__PURE__ */ Object.defineProperty({
  __proto__: null,
  default: Ki
}, Symbol.toStringTag, { value: "Module" })), Vi = /* @__PURE__ */ xi(Ji);
(function(t) {
  var r = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
  function e(P) {
    for (var G = "", U = 0, Q = 0, M = 0, X = 0, de = 0, le = 0, ne = 0, ye = 0; ye < P.length; )
      U = P.charCodeAt(ye++), X = U >> 2, Q = P.charCodeAt(ye++), de = (U & 3) << 4 | Q >> 4, M = P.charCodeAt(ye++), le = (Q & 15) << 2 | M >> 6, ne = M & 63, isNaN(Q) ? le = ne = 64 : isNaN(M) && (ne = 64), G += r.charAt(X) + r.charAt(de) + r.charAt(le) + r.charAt(ne);
    return G;
  }
  function n(P) {
    var G = "", U = 0, Q = 0, M = 0, X = 0, de = 0, le = 0, ne = 0;
    P = P.replace(/[^\w\+\/\=]/g, "");
    for (var ye = 0; ye < P.length; )
      X = r.indexOf(P.charAt(ye++)), de = r.indexOf(P.charAt(ye++)), U = X << 2 | de >> 4, G += String.fromCharCode(U), le = r.indexOf(P.charAt(ye++)), Q = (de & 15) << 4 | le >> 2, le !== 64 && (G += String.fromCharCode(Q)), ne = r.indexOf(P.charAt(ye++)), M = (le & 3) << 6 | ne, ne !== 64 && (G += String.fromCharCode(M));
    return G;
  }
  var i = function() {
    return typeof Buffer < "u" && typeof process < "u" && typeof process.versions < "u" && !!process.versions.node;
  }(), a = function() {
    if (typeof Buffer < "u") {
      var P = !Buffer.from;
      if (!P) try {
        Buffer.from("foo", "utf8");
      } catch {
        P = !0;
      }
      return P ? function(G, U) {
        return U ? new Buffer(G, U) : new Buffer(G);
      } : Buffer.from.bind(Buffer);
    }
    return function() {
    };
  }();
  function s(P) {
    if (i) {
      if (Buffer.alloc) return Buffer.alloc(P);
      var G = new Buffer(P);
      return G.fill(0), G;
    }
    return typeof Uint8Array < "u" ? new Uint8Array(P) : new Array(P);
  }
  function o(P) {
    return i ? Buffer.allocUnsafe ? Buffer.allocUnsafe(P) : new Buffer(P) : typeof Uint8Array < "u" ? new Uint8Array(P) : new Array(P);
  }
  var h = function(G) {
    return i ? a(G, "binary") : G.split("").map(function(U) {
      return U.charCodeAt(0) & 255;
    });
  }, l = /\u0000/g, f = /[\u0001-\u0006]/g, d = function(P) {
    for (var G = [], U = 0; U < P[0].length; ++U)
      G.push.apply(G, P[0][U]);
    return G;
  }, u = d, c = function(P, G, U) {
    for (var Q = [], M = G; M < U; M += 2) Q.push(String.fromCharCode(x(P, M)));
    return Q.join("").replace(l, "");
  }, g = c, p = function(P, G, U) {
    for (var Q = [], M = G; M < G + U; ++M) Q.push(("0" + P[M].toString(16)).slice(-2));
    return Q.join("");
  }, y = p, m = function(P) {
    if (Array.isArray(P[0])) return [].concat.apply([], P);
    var G = 0, U = 0;
    for (U = 0; U < P.length; ++U) G += P[U].length;
    var Q = new Uint8Array(G);
    for (U = 0, G = 0; U < P.length; G += P[U].length, ++U) Q.set(P[U], G);
    return Q;
  }, A = m;
  i && (c = function(P, G, U) {
    return Buffer.isBuffer(P) ? P.toString("utf16le", G, U).replace(l, "") : g(P, G, U);
  }, p = function(P, G, U) {
    return Buffer.isBuffer(P) ? P.toString("hex", G, G + U) : y(P, G, U);
  }, d = function(P) {
    return P[0].length > 0 && Buffer.isBuffer(P[0][0]) ? Buffer.concat(P[0]) : u(P);
  }, h = function(P) {
    return a(P, "binary");
  }, A = function(P) {
    return Buffer.isBuffer(P[0]) ? Buffer.concat(P) : m(P);
  });
  var E = function(P, G) {
    return P[G];
  }, x = function(P, G) {
    return P[G + 1] * 256 + P[G];
  }, S = function(P, G) {
    var U = P[G + 1] * 256 + P[G];
    return U < 32768 ? U : (65535 - U + 1) * -1;
  }, R = function(P, G) {
    return P[G + 3] * (1 << 24) + (P[G + 2] << 16) + (P[G + 1] << 8) + P[G];
  }, v = function(P, G) {
    return (P[G + 3] << 24) + (P[G + 2] << 16) + (P[G + 1] << 8) + P[G];
  };
  function O(P, G) {
    var U, Q, M = 0;
    switch (P) {
      case 1:
        U = E(this, this.l);
        break;
      case 2:
        U = (G !== "i" ? x : S)(this, this.l);
        break;
      case 4:
        U = v(this, this.l);
        break;
      case 16:
        M = 2, Q = p(this, this.l, P);
    }
    return this.l += P, M === 0 ? U : Q;
  }
  var C = function(P, G, U) {
    P[U] = G & 255, P[U + 1] = G >>> 8 & 255, P[U + 2] = G >>> 16 & 255, P[U + 3] = G >>> 24 & 255;
  }, B = function(P, G, U) {
    P[U] = G & 255, P[U + 1] = G >> 8 & 255, P[U + 2] = G >> 16 & 255, P[U + 3] = G >> 24 & 255;
  };
  function Z(P, G, U) {
    var Q = 0, M = 0;
    switch (U) {
      case "hex":
        for (; M < P; ++M)
          this[this.l++] = parseInt(G.slice(2 * M, 2 * M + 2), 16) || 0;
        return this;
      case "utf16le":
        var X = this.l + P;
        for (M = 0; M < Math.min(G.length, P); ++M) {
          var de = G.charCodeAt(M);
          this[this.l++] = de & 255, this[this.l++] = de >> 8;
        }
        for (; this.l < X; ) this[this.l++] = 0;
        return this;
    }
    switch (P) {
      case 1:
        Q = 1, this[this.l] = G & 255;
        break;
      case 2:
        Q = 2, this[this.l] = G & 255, G >>>= 8, this[this.l + 1] = G & 255;
        break;
      case 4:
        Q = 4, C(this, G, this.l);
        break;
      case -4:
        Q = 4, B(this, G, this.l);
        break;
    }
    return this.l += Q, this;
  }
  function T(P, G) {
    var U = p(this, this.l, P.length >> 1);
    if (U !== P) throw new Error(G + "Expected " + P + " saw " + U);
    this.l += P.length >> 1;
  }
  function j(P, G) {
    P.l = G, P.read_shift = O, P.chk = T, P.write_shift = Z;
  }
  function _(P) {
    var G = s(P);
    return j(G, 0), G;
  }
  /*! crc32.js (C) 2014-present SheetJS -- http://sheetjs.com */
  var K = function() {
    var P = {};
    P.version = "1.2.1";
    function G() {
      for (var W = 0, ie = new Array(256), $ = 0; $ != 256; ++$)
        W = $, W = W & 1 ? -306674912 ^ W >>> 1 : W >>> 1, W = W & 1 ? -306674912 ^ W >>> 1 : W >>> 1, W = W & 1 ? -306674912 ^ W >>> 1 : W >>> 1, W = W & 1 ? -306674912 ^ W >>> 1 : W >>> 1, W = W & 1 ? -306674912 ^ W >>> 1 : W >>> 1, W = W & 1 ? -306674912 ^ W >>> 1 : W >>> 1, W = W & 1 ? -306674912 ^ W >>> 1 : W >>> 1, W = W & 1 ? -306674912 ^ W >>> 1 : W >>> 1, ie[$] = W;
      return typeof Int32Array < "u" ? new Int32Array(ie) : ie;
    }
    var U = G();
    function Q(W) {
      var ie = 0, $ = 0, re = 0, he = typeof Int32Array < "u" ? new Int32Array(4096) : new Array(4096);
      for (re = 0; re != 256; ++re) he[re] = W[re];
      for (re = 0; re != 256; ++re)
        for ($ = W[re], ie = 256 + re; ie < 4096; ie += 256) $ = he[ie] = $ >>> 8 ^ W[$ & 255];
      var oe = [];
      for (re = 1; re != 16; ++re) oe[re - 1] = typeof Int32Array < "u" ? he.subarray(re * 256, re * 256 + 256) : he.slice(re * 256, re * 256 + 256);
      return oe;
    }
    var M = Q(U), X = M[0], de = M[1], le = M[2], ne = M[3], ye = M[4], Te = M[5], _e = M[6], Se = M[7], ve = M[8], Re = M[9], Ue = M[10], He = M[11], w = M[12], ae = M[13], ee = M[14];
    function D(W, ie) {
      for (var $ = ie ^ -1, re = 0, he = W.length; re < he; ) $ = $ >>> 8 ^ U[($ ^ W.charCodeAt(re++)) & 255];
      return ~$;
    }
    function F(W, ie) {
      for (var $ = ie ^ -1, re = W.length - 15, he = 0; he < re; ) $ = ee[W[he++] ^ $ & 255] ^ ae[W[he++] ^ $ >> 8 & 255] ^ w[W[he++] ^ $ >> 16 & 255] ^ He[W[he++] ^ $ >>> 24] ^ Ue[W[he++]] ^ Re[W[he++]] ^ ve[W[he++]] ^ Se[W[he++]] ^ _e[W[he++]] ^ Te[W[he++]] ^ ye[W[he++]] ^ ne[W[he++]] ^ le[W[he++]] ^ de[W[he++]] ^ X[W[he++]] ^ U[W[he++]];
      for (re += 15; he < re; ) $ = $ >>> 8 ^ U[($ ^ W[he++]) & 255];
      return ~$;
    }
    function z(W, ie) {
      for (var $ = ie ^ -1, re = 0, he = W.length, oe = 0, ge = 0; re < he; )
        oe = W.charCodeAt(re++), oe < 128 ? $ = $ >>> 8 ^ U[($ ^ oe) & 255] : oe < 2048 ? ($ = $ >>> 8 ^ U[($ ^ (192 | oe >> 6 & 31)) & 255], $ = $ >>> 8 ^ U[($ ^ (128 | oe & 63)) & 255]) : oe >= 55296 && oe < 57344 ? (oe = (oe & 1023) + 64, ge = W.charCodeAt(re++) & 1023, $ = $ >>> 8 ^ U[($ ^ (240 | oe >> 8 & 7)) & 255], $ = $ >>> 8 ^ U[($ ^ (128 | oe >> 2 & 63)) & 255], $ = $ >>> 8 ^ U[($ ^ (128 | ge >> 6 & 15 | (oe & 3) << 4)) & 255], $ = $ >>> 8 ^ U[($ ^ (128 | ge & 63)) & 255]) : ($ = $ >>> 8 ^ U[($ ^ (224 | oe >> 12 & 15)) & 255], $ = $ >>> 8 ^ U[($ ^ (128 | oe >> 6 & 63)) & 255], $ = $ >>> 8 ^ U[($ ^ (128 | oe & 63)) & 255]);
      return ~$;
    }
    return P.table = U, P.bstr = D, P.buf = F, P.str = z, P;
  }(), se = function() {
    var G = {};
    G.version = "1.2.2";
    function U(b, k) {
      for (var I = b.split("/"), N = k.split("/"), L = 0, H = 0, V = Math.min(I.length, N.length); L < V; ++L) {
        if (H = I[L].length - N[L].length) return H;
        if (I[L] != N[L]) return I[L] < N[L] ? -1 : 1;
      }
      return I.length - N.length;
    }
    function Q(b) {
      if (b.charAt(b.length - 1) == "/") return b.slice(0, -1).indexOf("/") === -1 ? b : Q(b.slice(0, -1));
      var k = b.lastIndexOf("/");
      return k === -1 ? b : b.slice(0, k + 1);
    }
    function M(b) {
      if (b.charAt(b.length - 1) == "/") return M(b.slice(0, -1));
      var k = b.lastIndexOf("/");
      return k === -1 ? b : b.slice(k + 1);
    }
    function X(b, k) {
      typeof k == "string" && (k = new Date(k));
      var I = k.getHours();
      I = I << 6 | k.getMinutes(), I = I << 5 | k.getSeconds() >>> 1, b.write_shift(2, I);
      var N = k.getFullYear() - 1980;
      N = N << 4 | k.getMonth() + 1, N = N << 5 | k.getDate(), b.write_shift(2, N);
    }
    function de(b) {
      var k = b.read_shift(2) & 65535, I = b.read_shift(2) & 65535, N = /* @__PURE__ */ new Date(), L = I & 31;
      I >>>= 5;
      var H = I & 15;
      I >>>= 4, N.setMilliseconds(0), N.setFullYear(I + 1980), N.setMonth(H - 1), N.setDate(L);
      var V = k & 31;
      k >>>= 5;
      var te = k & 63;
      return k >>>= 6, N.setHours(k), N.setMinutes(te), N.setSeconds(V << 1), N;
    }
    function le(b) {
      j(b, 0);
      for (var k = {}, I = 0; b.l <= b.length - 4; ) {
        var N = b.read_shift(2), L = b.read_shift(2), H = b.l + L, V = {};
        switch (N) {
          case 21589:
            I = b.read_shift(1), I & 1 && (V.mtime = b.read_shift(4)), L > 5 && (I & 2 && (V.atime = b.read_shift(4)), I & 4 && (V.ctime = b.read_shift(4))), V.mtime && (V.mt = new Date(V.mtime * 1e3));
            break;
        }
        b.l = H, k[N] = V;
      }
      return k;
    }
    var ne;
    function ye() {
      return ne || (ne = Vi);
    }
    function Te(b, k) {
      if (b[0] == 80 && b[1] == 75) return Rn(b, k);
      if ((b[0] | 32) == 109 && (b[1] | 32) == 105) return pi(b, k);
      if (b.length < 512) throw new Error("CFB file size " + b.length + " < 512");
      var I = 3, N = 512, L = 0, H = 0, V = 0, te = 0, J = 0, q = [], Y = b.slice(0, 512);
      j(Y, 0);
      var fe = _e(Y);
      switch (I = fe[0], I) {
        case 3:
          N = 512;
          break;
        case 4:
          N = 4096;
          break;
        case 0:
          if (fe[1] == 0) return Rn(b, k);
        default:
          throw new Error("Major Version: Expected 3 or 4 saw " + I);
      }
      N !== 512 && (Y = b.slice(0, N), j(
        Y,
        28
        /* blob.l */
      ));
      var ue = b.slice(0, N);
      Se(Y, I);
      var we = Y.read_shift(4, "i");
      if (I === 3 && we !== 0) throw new Error("# Directory Sectors: Expected 0 saw " + we);
      Y.l += 4, V = Y.read_shift(4, "i"), Y.l += 4, Y.chk("00100000", "Mini Stream Cutoff Size: "), te = Y.read_shift(4, "i"), L = Y.read_shift(4, "i"), J = Y.read_shift(4, "i"), H = Y.read_shift(4, "i");
      for (var pe = -1, me = 0; me < 109 && (pe = Y.read_shift(4, "i"), !(pe < 0)); ++me)
        q[me] = pe;
      var Ae = ve(b, N);
      He(J, H, Ae, N, q);
      var Ce = ae(Ae, V, q, N);
      Ce[V].name = "!Directory", L > 0 && te !== ge && (Ce[te].name = "!MiniFAT"), Ce[q[0]].name = "!FAT", Ce.fat_addrs = q, Ce.ssz = N;
      var Ne = {}, Me = [], Ot = [], Ht = [];
      ee(V, Ce, Ae, Me, L, Ne, Ot, te), Re(Ot, Ht, Me), Me.shift();
      var Dt = {
        FileIndex: Ot,
        FullPaths: Ht
      };
      return k && k.raw && (Dt.raw = { header: ue, sectors: Ae }), Dt;
    }
    function _e(b) {
      if (b[b.l] == 80 && b[b.l + 1] == 75) return [0, 0];
      b.chk(Ie, "Header Signature: "), b.l += 16;
      var k = b.read_shift(2, "u");
      return [b.read_shift(2, "u"), k];
    }
    function Se(b, k) {
      var I = 9;
      switch (b.l += 2, I = b.read_shift(2)) {
        case 9:
          if (k != 3) throw new Error("Sector Shift: Expected 9 saw " + I);
          break;
        case 12:
          if (k != 4) throw new Error("Sector Shift: Expected 12 saw " + I);
          break;
        default:
          throw new Error("Sector Shift: Expected 9 or 12 saw " + I);
      }
      b.chk("0600", "Mini Sector Shift: "), b.chk("000000000000", "Reserved: ");
    }
    function ve(b, k) {
      for (var I = Math.ceil(b.length / k) - 1, N = [], L = 1; L < I; ++L) N[L - 1] = b.slice(L * k, (L + 1) * k);
      return N[I - 1] = b.slice(I * k), N;
    }
    function Re(b, k, I) {
      for (var N = 0, L = 0, H = 0, V = 0, te = 0, J = I.length, q = [], Y = []; N < J; ++N)
        q[N] = Y[N] = N, k[N] = I[N];
      for (; te < Y.length; ++te)
        N = Y[te], L = b[N].L, H = b[N].R, V = b[N].C, q[N] === N && (L !== -1 && q[L] !== L && (q[N] = q[L]), H !== -1 && q[H] !== H && (q[N] = q[H])), V !== -1 && (q[V] = N), L !== -1 && N != q[N] && (q[L] = q[N], Y.lastIndexOf(L) < te && Y.push(L)), H !== -1 && N != q[N] && (q[H] = q[N], Y.lastIndexOf(H) < te && Y.push(H));
      for (N = 1; N < J; ++N) q[N] === N && (H !== -1 && q[H] !== H ? q[N] = q[H] : L !== -1 && q[L] !== L && (q[N] = q[L]));
      for (N = 1; N < J; ++N)
        if (b[N].type !== 0) {
          if (te = N, te != q[te]) do
            te = q[te], k[N] = k[te] + "/" + k[N];
          while (te !== 0 && q[te] !== -1 && te != q[te]);
          q[N] = -1;
        }
      for (k[0] += "/", N = 1; N < J; ++N)
        b[N].type !== 2 && (k[N] += "/");
    }
    function Ue(b, k, I) {
      for (var N = b.start, L = b.size, H = [], V = N; I && L > 0 && V >= 0; )
        H.push(k.slice(V * oe, V * oe + oe)), L -= oe, V = v(I, V * 4);
      return H.length === 0 ? _(0) : A(H).slice(0, b.size);
    }
    function He(b, k, I, N, L) {
      var H = ge;
      if (b === ge) {
        if (k !== 0) throw new Error("DIFAT chain shorter than expected");
      } else if (b !== -1) {
        var V = I[b], te = (N >>> 2) - 1;
        if (!V) return;
        for (var J = 0; J < te && (H = v(V, J * 4)) !== ge; ++J)
          L.push(H);
        k >= 1 && He(v(V, N - 4), k - 1, I, N, L);
      }
    }
    function w(b, k, I, N, L) {
      var H = [], V = [];
      L || (L = []);
      var te = N - 1, J = 0, q = 0;
      for (J = k; J >= 0; ) {
        L[J] = !0, H[H.length] = J, V.push(b[J]);
        var Y = I[Math.floor(J * 4 / N)];
        if (q = J * 4 & te, N < 4 + q) throw new Error("FAT boundary crossed: " + J + " 4 " + N);
        if (!b[Y]) break;
        J = v(b[Y], q);
      }
      return { nodes: H, data: d([V]) };
    }
    function ae(b, k, I, N) {
      var L = b.length, H = [], V = [], te = [], J = [], q = N - 1, Y = 0, fe = 0, ue = 0, we = 0;
      for (Y = 0; Y < L; ++Y)
        if (te = [], ue = Y + k, ue >= L && (ue -= L), !V[ue]) {
          J = [];
          var pe = [];
          for (fe = ue; fe >= 0; ) {
            pe[fe] = !0, V[fe] = !0, te[te.length] = fe, J.push(b[fe]);
            var me = I[Math.floor(fe * 4 / N)];
            if (we = fe * 4 & q, N < 4 + we) throw new Error("FAT boundary crossed: " + fe + " 4 " + N);
            if (!b[me] || (fe = v(b[me], we), pe[fe])) break;
          }
          H[ue] = { nodes: te, data: d([J]) };
        }
      return H;
    }
    function ee(b, k, I, N, L, H, V, te) {
      for (var J = 0, q = N.length ? 2 : 0, Y = k[b].data, fe = 0, ue = 0, we; fe < Y.length; fe += 128) {
        var pe = Y.slice(fe, fe + 128);
        j(pe, 64), ue = pe.read_shift(2), we = c(pe, 0, ue - q), N.push(we);
        var me = {
          name: we,
          type: pe.read_shift(1),
          color: pe.read_shift(1),
          L: pe.read_shift(4, "i"),
          R: pe.read_shift(4, "i"),
          C: pe.read_shift(4, "i"),
          clsid: pe.read_shift(16),
          state: pe.read_shift(4, "i"),
          start: 0,
          size: 0
        }, Ae = pe.read_shift(2) + pe.read_shift(2) + pe.read_shift(2) + pe.read_shift(2);
        Ae !== 0 && (me.ct = D(pe, pe.l - 8));
        var Ce = pe.read_shift(2) + pe.read_shift(2) + pe.read_shift(2) + pe.read_shift(2);
        Ce !== 0 && (me.mt = D(pe, pe.l - 8)), me.start = pe.read_shift(4, "i"), me.size = pe.read_shift(4, "i"), me.size < 0 && me.start < 0 && (me.size = me.type = 0, me.start = ge, me.name = ""), me.type === 5 ? (J = me.start, L > 0 && J !== ge && (k[J].name = "!StreamData")) : me.size >= 4096 ? (me.storage = "fat", k[me.start] === void 0 && (k[me.start] = w(I, me.start, k.fat_addrs, k.ssz)), k[me.start].name = me.name, me.content = k[me.start].data.slice(0, me.size)) : (me.storage = "minifat", me.size < 0 ? me.size = 0 : J !== ge && me.start !== ge && k[J] && (me.content = Ue(me, k[J].data, (k[te] || {}).data))), me.content && j(me.content, 0), H[we] = me, V.push(me);
      }
    }
    function D(b, k) {
      return new Date((R(b, k + 4) / 1e7 * Math.pow(2, 32) + R(b, k) / 1e7 - 11644473600) * 1e3);
    }
    function F(b, k) {
      return ye(), Te(ne.readFileSync(b), k);
    }
    function z(b, k) {
      var I = k && k.type;
      switch (I || i && Buffer.isBuffer(b) && (I = "buffer"), I || "base64") {
        case "file":
          return F(b, k);
        case "base64":
          return Te(h(n(b)), k);
        case "binary":
          return Te(h(b), k);
      }
      return Te(b, k);
    }
    function W(b, k) {
      var I = k || {}, N = I.root || "Root Entry";
      if (b.FullPaths || (b.FullPaths = []), b.FileIndex || (b.FileIndex = []), b.FullPaths.length !== b.FileIndex.length) throw new Error("inconsistent CFB structure");
      b.FullPaths.length === 0 && (b.FullPaths[0] = N + "/", b.FileIndex[0] = { name: N, type: 5 }), I.CLSID && (b.FileIndex[0].clsid = I.CLSID), ie(b);
    }
    function ie(b) {
      var k = "Sh33tJ5";
      if (!se.find(b, "/" + k)) {
        var I = _(4);
        I[0] = 55, I[1] = I[3] = 50, I[2] = 54, b.FileIndex.push({ name: k, type: 2, content: I, size: 4, L: 69, R: 69, C: 69 }), b.FullPaths.push(b.FullPaths[0] + k), $(b);
      }
    }
    function $(b, k) {
      W(b);
      for (var I = !1, N = !1, L = b.FullPaths.length - 1; L >= 0; --L) {
        var H = b.FileIndex[L];
        switch (H.type) {
          case 0:
            N ? I = !0 : (b.FileIndex.pop(), b.FullPaths.pop());
            break;
          case 1:
          case 2:
          case 5:
            N = !0, isNaN(H.R * H.L * H.C) && (I = !0), H.R > -1 && H.L > -1 && H.R == H.L && (I = !0);
            break;
          default:
            I = !0;
            break;
        }
      }
      if (!(!I && !k)) {
        var V = new Date(1987, 1, 19), te = 0, J = Object.create ? /* @__PURE__ */ Object.create(null) : {}, q = [];
        for (L = 0; L < b.FullPaths.length; ++L)
          J[b.FullPaths[L]] = !0, b.FileIndex[L].type !== 0 && q.push([b.FullPaths[L], b.FileIndex[L]]);
        for (L = 0; L < q.length; ++L) {
          var Y = Q(q[L][0]);
          for (N = J[Y]; !N; ) {
            for (; Q(Y) && !J[Q(Y)]; ) Y = Q(Y);
            q.push([Y, {
              name: M(Y).replace("/", ""),
              type: 1,
              clsid: De,
              ct: V,
              mt: V,
              content: null
            }]), J[Y] = !0, Y = Q(q[L][0]), N = J[Y];
          }
        }
        for (q.sort(function(we, pe) {
          return U(we[0], pe[0]);
        }), b.FullPaths = [], b.FileIndex = [], L = 0; L < q.length; ++L)
          b.FullPaths[L] = q[L][0], b.FileIndex[L] = q[L][1];
        for (L = 0; L < q.length; ++L) {
          var fe = b.FileIndex[L], ue = b.FullPaths[L];
          if (fe.name = M(ue).replace("/", ""), fe.L = fe.R = fe.C = -(fe.color = 1), fe.size = fe.content ? fe.content.length : 0, fe.start = 0, fe.clsid = fe.clsid || De, L === 0)
            fe.C = q.length > 1 ? 1 : -1, fe.size = 0, fe.type = 5;
          else if (ue.slice(-1) == "/") {
            for (te = L + 1; te < q.length && Q(b.FullPaths[te]) != ue; ++te) ;
            for (fe.C = te >= q.length ? -1 : te, te = L + 1; te < q.length && Q(b.FullPaths[te]) != Q(ue); ++te) ;
            fe.R = te >= q.length ? -1 : te, fe.type = 1;
          } else
            Q(b.FullPaths[L + 1] || "") == Q(ue) && (fe.R = L + 1), fe.type = 2;
        }
      }
    }
    function re(b, k) {
      var I = k || {};
      if (I.fileType == "mad") return ui(b, I);
      switch ($(b), I.fileType) {
        case "zip":
          return oi(b, I);
      }
      var N = function(we) {
        for (var pe = 0, me = 0, Ae = 0; Ae < we.FileIndex.length; ++Ae) {
          var Ce = we.FileIndex[Ae];
          if (Ce.content) {
            var Ne = Ce.content.length;
            Ne > 0 && (Ne < 4096 ? pe += Ne + 63 >> 6 : me += Ne + 511 >> 9);
          }
        }
        for (var Me = we.FullPaths.length + 3 >> 2, Ot = pe + 7 >> 3, Ht = pe + 127 >> 7, Dt = Ot + me + Me + Ht, gt = Dt + 127 >> 7, Dr = gt <= 109 ? 0 : Math.ceil((gt - 109) / 127); Dt + gt + Dr + 127 >> 7 > gt; ) Dr = ++gt <= 109 ? 0 : Math.ceil((gt - 109) / 127);
        var at = [1, Dr, gt, Ht, Me, me, pe, 0];
        return we.FileIndex[0].size = pe << 6, at[7] = (we.FileIndex[0].start = at[0] + at[1] + at[2] + at[3] + at[4] + at[5]) + (at[6] + 7 >> 3), at;
      }(b), L = _(N[7] << 9), H = 0, V = 0;
      {
        for (H = 0; H < 8; ++H) L.write_shift(1, be[H]);
        for (H = 0; H < 8; ++H) L.write_shift(2, 0);
        for (L.write_shift(2, 62), L.write_shift(2, 3), L.write_shift(2, 65534), L.write_shift(2, 9), L.write_shift(2, 6), H = 0; H < 3; ++H) L.write_shift(2, 0);
        for (L.write_shift(4, 0), L.write_shift(4, N[2]), L.write_shift(4, N[0] + N[1] + N[2] + N[3] - 1), L.write_shift(4, 0), L.write_shift(4, 4096), L.write_shift(4, N[3] ? N[0] + N[1] + N[2] - 1 : ge), L.write_shift(4, N[3]), L.write_shift(-4, N[1] ? N[0] - 1 : ge), L.write_shift(4, N[1]), H = 0; H < 109; ++H) L.write_shift(-4, H < N[2] ? N[1] + H : -1);
      }
      if (N[1])
        for (V = 0; V < N[1]; ++V) {
          for (; H < 236 + V * 127; ++H) L.write_shift(-4, H < N[2] ? N[1] + H : -1);
          L.write_shift(-4, V === N[1] - 1 ? ge : V + 1);
        }
      var te = function(we) {
        for (V += we; H < V - 1; ++H) L.write_shift(-4, H + 1);
        we && (++H, L.write_shift(-4, ge));
      };
      for (V = H = 0, V += N[1]; H < V; ++H) L.write_shift(-4, Je.DIFSECT);
      for (V += N[2]; H < V; ++H) L.write_shift(-4, Je.FATSECT);
      te(N[3]), te(N[4]);
      for (var J = 0, q = 0, Y = b.FileIndex[0]; J < b.FileIndex.length; ++J)
        Y = b.FileIndex[J], Y.content && (q = Y.content.length, !(q < 4096) && (Y.start = V, te(q + 511 >> 9)));
      for (te(N[6] + 7 >> 3); L.l & 511; ) L.write_shift(-4, Je.ENDOFCHAIN);
      for (V = H = 0, J = 0; J < b.FileIndex.length; ++J)
        Y = b.FileIndex[J], Y.content && (q = Y.content.length, !(!q || q >= 4096) && (Y.start = V, te(q + 63 >> 6)));
      for (; L.l & 511; ) L.write_shift(-4, Je.ENDOFCHAIN);
      for (H = 0; H < N[4] << 2; ++H) {
        var fe = b.FullPaths[H];
        if (!fe || fe.length === 0) {
          for (J = 0; J < 17; ++J) L.write_shift(4, 0);
          for (J = 0; J < 3; ++J) L.write_shift(4, -1);
          for (J = 0; J < 12; ++J) L.write_shift(4, 0);
          continue;
        }
        Y = b.FileIndex[H], H === 0 && (Y.start = Y.size ? Y.start - 1 : ge);
        var ue = H === 0 && I.root || Y.name;
        if (ue.length > 32 && (console.error("Name " + ue + " will be truncated to " + ue.slice(0, 32)), ue = ue.slice(0, 32)), q = 2 * (ue.length + 1), L.write_shift(64, ue, "utf16le"), L.write_shift(2, q), L.write_shift(1, Y.type), L.write_shift(1, Y.color), L.write_shift(-4, Y.L), L.write_shift(-4, Y.R), L.write_shift(-4, Y.C), Y.clsid) L.write_shift(16, Y.clsid, "hex");
        else for (J = 0; J < 4; ++J) L.write_shift(4, 0);
        L.write_shift(4, Y.state || 0), L.write_shift(4, 0), L.write_shift(4, 0), L.write_shift(4, 0), L.write_shift(4, 0), L.write_shift(4, Y.start), L.write_shift(4, Y.size), L.write_shift(4, 0);
      }
      for (H = 1; H < b.FileIndex.length; ++H)
        if (Y = b.FileIndex[H], Y.size >= 4096)
          if (L.l = Y.start + 1 << 9, i && Buffer.isBuffer(Y.content))
            Y.content.copy(L, L.l, 0, Y.size), L.l += Y.size + 511 & -512;
          else {
            for (J = 0; J < Y.size; ++J) L.write_shift(1, Y.content[J]);
            for (; J & 511; ++J) L.write_shift(1, 0);
          }
      for (H = 1; H < b.FileIndex.length; ++H)
        if (Y = b.FileIndex[H], Y.size > 0 && Y.size < 4096)
          if (i && Buffer.isBuffer(Y.content))
            Y.content.copy(L, L.l, 0, Y.size), L.l += Y.size + 63 & -64;
          else {
            for (J = 0; J < Y.size; ++J) L.write_shift(1, Y.content[J]);
            for (; J & 63; ++J) L.write_shift(1, 0);
          }
      if (i)
        L.l = L.length;
      else
        for (; L.l < L.length; ) L.write_shift(1, 0);
      return L;
    }
    function he(b, k) {
      var I = b.FullPaths.map(function(J) {
        return J.toUpperCase();
      }), N = I.map(function(J) {
        var q = J.split("/");
        return q[q.length - (J.slice(-1) == "/" ? 2 : 1)];
      }), L = !1;
      k.charCodeAt(0) === 47 ? (L = !0, k = I[0].slice(0, -1) + k) : L = k.indexOf("/") !== -1;
      var H = k.toUpperCase(), V = L === !0 ? I.indexOf(H) : N.indexOf(H);
      if (V !== -1) return b.FileIndex[V];
      var te = !H.match(f);
      for (H = H.replace(l, ""), te && (H = H.replace(f, "!")), V = 0; V < I.length; ++V)
        if ((te ? I[V].replace(f, "!") : I[V]).replace(l, "") == H || (te ? N[V].replace(f, "!") : N[V]).replace(l, "") == H) return b.FileIndex[V];
      return null;
    }
    var oe = 64, ge = -2, Ie = "d0cf11e0a1b11ae1", be = [208, 207, 17, 224, 161, 177, 26, 225], De = "00000000000000000000000000000000", Je = {
      /* 2.1 Compund File Sector Numbers and Types */
      MAXREGSECT: -6,
      DIFSECT: -4,
      FATSECT: -3,
      ENDOFCHAIN: ge,
      FREESECT: -1,
      /* 2.2 Compound File Header */
      HEADER_SIGNATURE: Ie,
      HEADER_MINOR_VERSION: "3e00",
      MAXREGSID: -6,
      NOSTREAM: -1,
      HEADER_CLSID: De,
      /* 2.6.1 Compound File Directory Entry */
      EntryTypes: ["unknown", "storage", "stream", "lockbytes", "property", "root"]
    };
    function Xe(b, k, I) {
      ye();
      var N = re(b, I);
      ne.writeFileSync(k, N);
    }
    function Ve(b) {
      for (var k = new Array(b.length), I = 0; I < b.length; ++I) k[I] = String.fromCharCode(b[I]);
      return k.join("");
    }
    function Fe(b, k) {
      var I = re(b, k);
      switch (k && k.type || "buffer") {
        case "file":
          return ye(), ne.writeFileSync(k.filename, I), I;
        case "binary":
          return typeof I == "string" ? I : Ve(I);
        case "base64":
          return e(typeof I == "string" ? I : Ve(I));
        case "buffer":
          if (i) return Buffer.isBuffer(I) ? I : a(I);
        case "array":
          return typeof I == "string" ? h(I) : I;
      }
      return I;
    }
    var qe;
    function Bt(b) {
      try {
        var k = b.InflateRaw, I = new k();
        if (I._processChunk(new Uint8Array([3, 0]), I._finishFlushFlag), I.bytesRead) qe = b;
        else throw new Error("zlib does not expose bytesRead");
      } catch (N) {
        console.error("cannot use native zlib: " + (N.message || N));
      }
    }
    function Ye(b, k) {
      if (!qe) return Nn(b, k);
      var I = qe.InflateRaw, N = new I(), L = N._processChunk(b.slice(b.l), N._finishFlushFlag);
      return b.l += N.bytesRead, L;
    }
    function cr(b) {
      return qe ? qe.deflateRawSync(b) : bn(b);
    }
    var Pr = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15], Tt = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258], Et = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
    function fr(b) {
      var k = (b << 1 | b << 11) & 139536 | (b << 5 | b << 15) & 558144;
      return (k >> 16 | k >> 8 | k) & 255;
    }
    for (var ke = typeof Uint8Array < "u", ze = ke ? new Uint8Array(256) : [], dr = 0; dr < 256; ++dr) ze[dr] = fr(dr);
    function Qa(b, k) {
      var I = ze[b & 255];
      return k <= 8 ? I >>> 8 - k : (I = I << 8 | ze[b >> 8 & 255], k <= 16 ? I >>> 16 - k : (I = I << 8 | ze[b >> 16 & 255], I >>> 24 - k));
    }
    function ei(b, k) {
      var I = k & 7, N = k >>> 3;
      return (b[N] | (I <= 6 ? 0 : b[N + 1] << 8)) >>> I & 3;
    }
    function Fr(b, k) {
      var I = k & 7, N = k >>> 3;
      return (b[N] | (I <= 5 ? 0 : b[N + 1] << 8)) >>> I & 7;
    }
    function ti(b, k) {
      var I = k & 7, N = k >>> 3;
      return (b[N] | (I <= 4 ? 0 : b[N + 1] << 8)) >>> I & 15;
    }
    function _n(b, k) {
      var I = k & 7, N = k >>> 3;
      return (b[N] | (I <= 3 ? 0 : b[N + 1] << 8)) >>> I & 31;
    }
    function yn(b, k) {
      var I = k & 7, N = k >>> 3;
      return (b[N] | (I <= 1 ? 0 : b[N + 1] << 8)) >>> I & 127;
    }
    function pr(b, k, I) {
      var N = k & 7, L = k >>> 3, H = (1 << I) - 1, V = b[L] >>> N;
      return I < 8 - N || (V |= b[L + 1] << 8 - N, I < 16 - N) || (V |= b[L + 2] << 16 - N, I < 24 - N) || (V |= b[L + 3] << 24 - N), V & H;
    }
    function An(b, k, I) {
      var N = k & 7, L = k >>> 3;
      return N <= 5 ? b[L] |= (I & 7) << N : (b[L] |= I << N & 255, b[L + 1] = (I & 7) >> 8 - N), k + 3;
    }
    function ri(b, k, I) {
      var N = k & 7, L = k >>> 3;
      return I = (I & 1) << N, b[L] |= I, k + 1;
    }
    function It(b, k, I) {
      var N = k & 7, L = k >>> 3;
      return I <<= N, b[L] |= I & 255, I >>>= 8, b[L + 1] = I, k + 8;
    }
    function Sn(b, k, I) {
      var N = k & 7, L = k >>> 3;
      return I <<= N, b[L] |= I & 255, I >>>= 8, b[L + 1] = I & 255, b[L + 2] = I >>> 8, k + 16;
    }
    function kr(b, k) {
      var I = b.length, N = 2 * I > k ? 2 * I : k + 5, L = 0;
      if (I >= k) return b;
      if (i) {
        var H = o(N);
        if (b.copy) b.copy(H);
        else for (; L < b.length; ++L) H[L] = b[L];
        return H;
      } else if (ke) {
        var V = new Uint8Array(N);
        if (V.set) V.set(b);
        else for (; L < I; ++L) V[L] = b[L];
        return V;
      }
      return b.length = N, b;
    }
    function Qe(b) {
      for (var k = new Array(b), I = 0; I < b; ++I) k[I] = 0;
      return k;
    }
    function ur(b, k, I) {
      var N = 1, L = 0, H = 0, V = 0, te = 0, J = b.length, q = ke ? new Uint16Array(32) : Qe(32);
      for (H = 0; H < 32; ++H) q[H] = 0;
      for (H = J; H < I; ++H) b[H] = 0;
      J = b.length;
      var Y = ke ? new Uint16Array(J) : Qe(J);
      for (H = 0; H < J; ++H)
        q[L = b[H]]++, N < L && (N = L), Y[H] = 0;
      for (q[0] = 0, H = 1; H <= N; ++H) q[H + 16] = te = te + q[H - 1] << 1;
      for (H = 0; H < J; ++H)
        te = b[H], te != 0 && (Y[H] = q[te + 16]++);
      var fe = 0;
      for (H = 0; H < J; ++H)
        if (fe = b[H], fe != 0)
          for (te = Qa(Y[H], N) >> N - fe, V = (1 << N + 4 - fe) - 1; V >= 0; --V)
            k[te | V << fe] = fe & 15 | H << 4;
      return N;
    }
    var Lr = ke ? new Uint16Array(512) : Qe(512), Br = ke ? new Uint16Array(32) : Qe(32);
    if (!ke) {
      for (var ut = 0; ut < 512; ++ut) Lr[ut] = 0;
      for (ut = 0; ut < 32; ++ut) Br[ut] = 0;
    }
    (function() {
      for (var b = [], k = 0; k < 32; k++) b.push(5);
      ur(b, Br, 32);
      var I = [];
      for (k = 0; k <= 143; k++) I.push(8);
      for (; k <= 255; k++) I.push(9);
      for (; k <= 279; k++) I.push(7);
      for (; k <= 287; k++) I.push(8);
      ur(I, Lr, 288);
    })();
    var ni = function() {
      for (var k = ke ? new Uint8Array(32768) : [], I = 0, N = 0; I < Et.length - 1; ++I)
        for (; N < Et[I + 1]; ++N) k[N] = I;
      for (; N < 32768; ++N) k[N] = 29;
      var L = ke ? new Uint8Array(259) : [];
      for (I = 0, N = 0; I < Tt.length - 1; ++I)
        for (; N < Tt[I + 1]; ++N) L[N] = I;
      function H(te, J) {
        for (var q = 0; q < te.length; ) {
          var Y = Math.min(65535, te.length - q), fe = q + Y == te.length;
          for (J.write_shift(1, +fe), J.write_shift(2, Y), J.write_shift(2, ~Y & 65535); Y-- > 0; ) J[J.l++] = te[q++];
        }
        return J.l;
      }
      function V(te, J) {
        for (var q = 0, Y = 0, fe = ke ? new Uint16Array(32768) : []; Y < te.length; ) {
          var ue = (
            /* data.length - boff; */
            Math.min(65535, te.length - Y)
          );
          if (ue < 10) {
            for (q = An(J, q, +(Y + ue == te.length)), q & 7 && (q += 8 - (q & 7)), J.l = q / 8 | 0, J.write_shift(2, ue), J.write_shift(2, ~ue & 65535); ue-- > 0; ) J[J.l++] = te[Y++];
            q = J.l * 8;
            continue;
          }
          q = An(J, q, +(Y + ue == te.length) + 2);
          for (var we = 0; ue-- > 0; ) {
            var pe = te[Y];
            we = (we << 5 ^ pe) & 32767;
            var me = -1, Ae = 0;
            if ((me = fe[we]) && (me |= Y & -32768, me > Y && (me -= 32768), me < Y))
              for (; te[me + Ae] == te[Y + Ae] && Ae < 250; ) ++Ae;
            if (Ae > 2) {
              pe = L[Ae], pe <= 22 ? q = It(J, q, ze[pe + 1] >> 1) - 1 : (It(J, q, 3), q += 5, It(J, q, ze[pe - 23] >> 5), q += 3);
              var Ce = pe < 8 ? 0 : pe - 4 >> 2;
              Ce > 0 && (Sn(J, q, Ae - Tt[pe]), q += Ce), pe = k[Y - me], q = It(J, q, ze[pe] >> 3), q -= 3;
              var Ne = pe < 4 ? 0 : pe - 2 >> 1;
              Ne > 0 && (Sn(J, q, Y - me - Et[pe]), q += Ne);
              for (var Me = 0; Me < Ae; ++Me)
                fe[we] = Y & 32767, we = (we << 5 ^ te[Y]) & 32767, ++Y;
              ue -= Ae - 1;
            } else
              pe <= 143 ? pe = pe + 48 : q = ri(J, q, 1), q = It(J, q, ze[pe]), fe[we] = Y & 32767, ++Y;
          }
          q = It(J, q, 0) - 1;
        }
        return J.l = (q + 7) / 8 | 0, J.l;
      }
      return function(J, q) {
        return J.length < 8 ? H(J, q) : V(J, q);
      };
    }();
    function bn(b) {
      var k = _(50 + Math.floor(b.length * 1.1)), I = ni(b, k);
      return k.slice(0, I);
    }
    var xn = ke ? new Uint16Array(32768) : Qe(32768), Tn = ke ? new Uint16Array(32768) : Qe(32768), En = ke ? new Uint16Array(128) : Qe(128), In = 1, vn = 1;
    function ai(b, k) {
      var I = _n(b, k) + 257;
      k += 5;
      var N = _n(b, k) + 1;
      k += 5;
      var L = ti(b, k) + 4;
      k += 4;
      for (var H = 0, V = ke ? new Uint8Array(19) : Qe(19), te = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], J = 1, q = ke ? new Uint8Array(8) : Qe(8), Y = ke ? new Uint8Array(8) : Qe(8), fe = V.length, ue = 0; ue < L; ++ue)
        V[Pr[ue]] = H = Fr(b, k), J < H && (J = H), q[H]++, k += 3;
      var we = 0;
      for (q[0] = 0, ue = 1; ue <= J; ++ue) Y[ue] = we = we + q[ue - 1] << 1;
      for (ue = 0; ue < fe; ++ue) (we = V[ue]) != 0 && (te[ue] = Y[we]++);
      var pe = 0;
      for (ue = 0; ue < fe; ++ue)
        if (pe = V[ue], pe != 0) {
          we = ze[te[ue]] >> 8 - pe;
          for (var me = (1 << 7 - pe) - 1; me >= 0; --me) En[we | me << pe] = pe & 7 | ue << 3;
        }
      var Ae = [];
      for (J = 1; Ae.length < I + N; )
        switch (we = En[yn(b, k)], k += we & 7, we >>>= 3) {
          case 16:
            for (H = 3 + ei(b, k), k += 2, we = Ae[Ae.length - 1]; H-- > 0; ) Ae.push(we);
            break;
          case 17:
            for (H = 3 + Fr(b, k), k += 3; H-- > 0; ) Ae.push(0);
            break;
          case 18:
            for (H = 11 + yn(b, k), k += 7; H-- > 0; ) Ae.push(0);
            break;
          default:
            Ae.push(we), J < we && (J = we);
            break;
        }
      var Ce = Ae.slice(0, I), Ne = Ae.slice(I);
      for (ue = I; ue < 286; ++ue) Ce[ue] = 0;
      for (ue = N; ue < 30; ++ue) Ne[ue] = 0;
      return In = ur(Ce, xn, 286), vn = ur(Ne, Tn, 30), k;
    }
    function ii(b, k) {
      if (b[0] == 3 && !(b[1] & 3))
        return [s(k), 2];
      for (var I = 0, N = 0, L = o(k || 1 << 18), H = 0, V = L.length >>> 0, te = 0, J = 0; !(N & 1); ) {
        if (N = Fr(b, I), I += 3, N >>> 1)
          N >> 1 == 1 ? (te = 9, J = 5) : (I = ai(b, I), te = In, J = vn);
        else {
          I & 7 && (I += 8 - (I & 7));
          var q = b[I >>> 3] | b[(I >>> 3) + 1] << 8;
          if (I += 32, q > 0)
            for (!k && V < H + q && (L = kr(L, H + q), V = L.length); q-- > 0; )
              L[H++] = b[I >>> 3], I += 8;
          continue;
        }
        for (; ; ) {
          !k && V < H + 32767 && (L = kr(L, H + 32767), V = L.length);
          var Y = pr(b, I, te), fe = N >>> 1 == 1 ? Lr[Y] : xn[Y];
          if (I += fe & 15, fe >>>= 4, !(fe >>> 8 & 255)) L[H++] = fe;
          else {
            if (fe == 256) break;
            fe -= 257;
            var ue = fe < 8 ? 0 : fe - 4 >> 2;
            ue > 5 && (ue = 0);
            var we = H + Tt[fe];
            ue > 0 && (we += pr(b, I, ue), I += ue), Y = pr(b, I, J), fe = N >>> 1 == 1 ? Br[Y] : Tn[Y], I += fe & 15, fe >>>= 4;
            var pe = fe < 4 ? 0 : fe - 2 >> 1, me = Et[fe];
            for (pe > 0 && (me += pr(b, I, pe), I += pe), !k && V < we && (L = kr(L, we + 100), V = L.length); H < we; )
              L[H] = L[H - me], ++H;
          }
        }
      }
      return k ? [L, I + 7 >>> 3] : [L.slice(0, H), I + 7 >>> 3];
    }
    function Nn(b, k) {
      var I = b.slice(b.l || 0), N = ii(I, k);
      return b.l += N[1], N[0];
    }
    function Or(b, k) {
      if (b)
        typeof console < "u" && console.error(k);
      else throw new Error(k);
    }
    function Rn(b, k) {
      var I = b;
      j(I, 0);
      var N = [], L = [], H = {
        FileIndex: N,
        FullPaths: L
      };
      W(H, { root: k.root });
      for (var V = I.length - 4; (I[V] != 80 || I[V + 1] != 75 || I[V + 2] != 5 || I[V + 3] != 6) && V >= 0; ) --V;
      I.l = V + 4, I.l += 4;
      var te = I.read_shift(2);
      I.l += 6;
      var J = I.read_shift(4);
      for (I.l = J, V = 0; V < te; ++V) {
        I.l += 20;
        var q = I.read_shift(4), Y = I.read_shift(4), fe = I.read_shift(2), ue = I.read_shift(2), we = I.read_shift(2);
        I.l += 8;
        var pe = I.read_shift(4), me = le(I.slice(I.l + fe, I.l + fe + ue));
        I.l += fe + ue + we;
        var Ae = I.l;
        I.l = pe + 4, si(I, q, Y, H, me), I.l = Ae;
      }
      return H;
    }
    function si(b, k, I, N, L) {
      b.l += 2;
      var H = b.read_shift(2), V = b.read_shift(2), te = de(b);
      if (H & 8257) throw new Error("Unsupported ZIP encryption");
      for (var J = b.read_shift(4), q = b.read_shift(4), Y = b.read_shift(4), fe = b.read_shift(2), ue = b.read_shift(2), we = "", pe = 0; pe < fe; ++pe) we += String.fromCharCode(b[b.l++]);
      if (ue) {
        var me = le(b.slice(b.l, b.l + ue));
        (me[21589] || {}).mt && (te = me[21589].mt), ((L || {})[21589] || {}).mt && (te = L[21589].mt);
      }
      b.l += ue;
      var Ae = b.slice(b.l, b.l + q);
      switch (V) {
        case 8:
          Ae = Ye(b, Y);
          break;
        case 0:
          break;
        default:
          throw new Error("Unsupported ZIP Compression method " + V);
      }
      var Ce = !1;
      H & 8 && (J = b.read_shift(4), J == 134695760 && (J = b.read_shift(4), Ce = !0), q = b.read_shift(4), Y = b.read_shift(4)), q != k && Or(Ce, "Bad compressed size: " + k + " != " + q), Y != I && Or(Ce, "Bad uncompressed size: " + I + " != " + Y);
      var Ne = K.buf(Ae, 0);
      J >> 0 != Ne >> 0 && Or(Ce, "Bad CRC32 checksum: " + J + " != " + Ne), Hr(N, we, Ae, { unsafe: !0, mt: te });
    }
    function oi(b, k) {
      var I = k || {}, N = [], L = [], H = _(1), V = I.compression ? 8 : 0, te = 0, J = 0, q = 0, Y = 0, fe = 0, ue = b.FullPaths[0], we = ue, pe = b.FileIndex[0], me = [], Ae = 0;
      for (J = 1; J < b.FullPaths.length; ++J)
        if (we = b.FullPaths[J].slice(ue.length), pe = b.FileIndex[J], !(!pe.size || !pe.content || we == "Sh33tJ5")) {
          var Ce = Y, Ne = _(we.length);
          for (q = 0; q < we.length; ++q) Ne.write_shift(1, we.charCodeAt(q) & 127);
          Ne = Ne.slice(0, Ne.l), me[fe] = K.buf(pe.content, 0);
          var Me = pe.content;
          V == 8 && (Me = cr(Me)), H = _(30), H.write_shift(4, 67324752), H.write_shift(2, 20), H.write_shift(2, te), H.write_shift(2, V), pe.mt ? X(H, pe.mt) : H.write_shift(4, 0), H.write_shift(-4, me[fe]), H.write_shift(4, Me.length), H.write_shift(4, pe.content.length), H.write_shift(2, Ne.length), H.write_shift(2, 0), Y += H.length, N.push(H), Y += Ne.length, N.push(Ne), Y += Me.length, N.push(Me), H = _(46), H.write_shift(4, 33639248), H.write_shift(2, 0), H.write_shift(2, 20), H.write_shift(2, te), H.write_shift(2, V), H.write_shift(4, 0), H.write_shift(-4, me[fe]), H.write_shift(4, Me.length), H.write_shift(4, pe.content.length), H.write_shift(2, Ne.length), H.write_shift(2, 0), H.write_shift(2, 0), H.write_shift(2, 0), H.write_shift(2, 0), H.write_shift(4, 0), H.write_shift(4, Ce), Ae += H.l, L.push(H), Ae += Ne.length, L.push(Ne), ++fe;
        }
      return H = _(22), H.write_shift(4, 101010256), H.write_shift(2, 0), H.write_shift(2, 0), H.write_shift(2, fe), H.write_shift(2, fe), H.write_shift(4, Ae), H.write_shift(4, Y), H.write_shift(2, 0), A([A(N), A(L), H]);
    }
    var gr = {
      htm: "text/html",
      xml: "text/xml",
      gif: "image/gif",
      jpg: "image/jpeg",
      png: "image/png",
      mso: "application/x-mso",
      thmx: "application/vnd.ms-officetheme",
      sh33tj5: "application/octet-stream"
    };
    function li(b, k) {
      if (b.ctype) return b.ctype;
      var I = b.name || "", N = I.match(/\.([^\.]+)$/);
      return N && gr[N[1]] || k && (N = (I = k).match(/[\.\\]([^\.\\])+$/), N && gr[N[1]]) ? gr[N[1]] : "application/octet-stream";
    }
    function hi(b) {
      for (var k = e(b), I = [], N = 0; N < k.length; N += 76) I.push(k.slice(N, N + 76));
      return I.join(`\r
`) + `\r
`;
    }
    function ci(b) {
      var k = b.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7E-\xFF=]/g, function(q) {
        var Y = q.charCodeAt(0).toString(16).toUpperCase();
        return "=" + (Y.length == 1 ? "0" + Y : Y);
      });
      k = k.replace(/ $/mg, "=20").replace(/\t$/mg, "=09"), k.charAt(0) == `
` && (k = "=0D" + k.slice(1)), k = k.replace(/\r(?!\n)/mg, "=0D").replace(/\n\n/mg, `
=0A`).replace(/([^\r\n])\n/mg, "$1=0A");
      for (var I = [], N = k.split(`\r
`), L = 0; L < N.length; ++L) {
        var H = N[L];
        if (H.length == 0) {
          I.push("");
          continue;
        }
        for (var V = 0; V < H.length; ) {
          var te = 76, J = H.slice(V, V + te);
          J.charAt(te - 1) == "=" ? te-- : J.charAt(te - 2) == "=" ? te -= 2 : J.charAt(te - 3) == "=" && (te -= 3), J = H.slice(V, V + te), V += te, V < H.length && (J += "="), I.push(J);
        }
      }
      return I.join(`\r
`);
    }
    function fi(b) {
      for (var k = [], I = 0; I < b.length; ++I) {
        for (var N = b[I]; I <= b.length && N.charAt(N.length - 1) == "="; ) N = N.slice(0, N.length - 1) + b[++I];
        k.push(N);
      }
      for (var L = 0; L < k.length; ++L) k[L] = k[L].replace(/[=][0-9A-Fa-f]{2}/g, function(H) {
        return String.fromCharCode(parseInt(H.slice(1), 16));
      });
      return h(k.join(`\r
`));
    }
    function di(b, k, I) {
      for (var N = "", L = "", H = "", V, te = 0; te < 10; ++te) {
        var J = k[te];
        if (!J || J.match(/^\s*$/)) break;
        var q = J.match(/^(.*?):\s*([^\s].*)$/);
        if (q) switch (q[1].toLowerCase()) {
          case "content-location":
            N = q[2].trim();
            break;
          case "content-type":
            H = q[2].trim();
            break;
          case "content-transfer-encoding":
            L = q[2].trim();
            break;
        }
      }
      switch (++te, L.toLowerCase()) {
        case "base64":
          V = h(n(k.slice(te).join("")));
          break;
        case "quoted-printable":
          V = fi(k.slice(te));
          break;
        default:
          throw new Error("Unsupported Content-Transfer-Encoding " + L);
      }
      var Y = Hr(b, N.slice(I.length), V, { unsafe: !0 });
      H && (Y.ctype = H);
    }
    function pi(b, k) {
      if (Ve(b.slice(0, 13)).toLowerCase() != "mime-version:") throw new Error("Unsupported MAD header");
      var I = k && k.root || "", N = (i && Buffer.isBuffer(b) ? b.toString("binary") : Ve(b)).split(`\r
`), L = 0, H = "";
      for (L = 0; L < N.length; ++L)
        if (H = N[L], !!/^Content-Location:/i.test(H) && (H = H.slice(H.indexOf("file")), I || (I = H.slice(0, H.lastIndexOf("/") + 1)), H.slice(0, I.length) != I))
          for (; I.length > 0 && (I = I.slice(0, I.length - 1), I = I.slice(0, I.lastIndexOf("/") + 1), H.slice(0, I.length) != I); )
            ;
      var V = (N[1] || "").match(/boundary="(.*?)"/);
      if (!V) throw new Error("MAD cannot find boundary");
      var te = "--" + (V[1] || ""), J = [], q = [], Y = {
        FileIndex: J,
        FullPaths: q
      };
      W(Y);
      var fe, ue = 0;
      for (L = 0; L < N.length; ++L) {
        var we = N[L];
        we !== te && we !== te + "--" || (ue++ && di(Y, N.slice(fe, L), I), fe = L);
      }
      return Y;
    }
    function ui(b, k) {
      var I = k || {}, N = I.boundary || "SheetJS";
      N = "------=" + N;
      for (var L = [
        "MIME-Version: 1.0",
        'Content-Type: multipart/related; boundary="' + N.slice(2) + '"',
        "",
        "",
        ""
      ], H = b.FullPaths[0], V = H, te = b.FileIndex[0], J = 1; J < b.FullPaths.length; ++J)
        if (V = b.FullPaths[J].slice(H.length), te = b.FileIndex[J], !(!te.size || !te.content || V == "Sh33tJ5")) {
          V = V.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7E-\xFF]/g, function(Ae) {
            return "_x" + Ae.charCodeAt(0).toString(16) + "_";
          }).replace(/[\u0080-\uFFFF]/g, function(Ae) {
            return "_u" + Ae.charCodeAt(0).toString(16) + "_";
          });
          for (var q = te.content, Y = i && Buffer.isBuffer(q) ? q.toString("binary") : Ve(q), fe = 0, ue = Math.min(1024, Y.length), we = 0, pe = 0; pe <= ue; ++pe) (we = Y.charCodeAt(pe)) >= 32 && we < 128 && ++fe;
          var me = fe >= ue * 4 / 5;
          L.push(N), L.push("Content-Location: " + (I.root || "file:///C:/SheetJS/") + V), L.push("Content-Transfer-Encoding: " + (me ? "quoted-printable" : "base64")), L.push("Content-Type: " + li(te, V)), L.push(""), L.push(me ? ci(Y) : hi(Y));
        }
      return L.push(N + `--\r
`), L.join(`\r
`);
    }
    function gi(b) {
      var k = {};
      return W(k, b), k;
    }
    function Hr(b, k, I, N) {
      var L = N && N.unsafe;
      L || W(b);
      var H = !L && se.find(b, k);
      if (!H) {
        var V = b.FullPaths[0];
        k.slice(0, V.length) == V ? V = k : (V.slice(-1) != "/" && (V += "/"), V = (V + k).replace("//", "/")), H = { name: M(k), type: 2 }, b.FileIndex.push(H), b.FullPaths.push(V), L || se.utils.cfb_gc(b);
      }
      return H.content = I, H.size = I ? I.length : 0, N && (N.CLSID && (H.clsid = N.CLSID), N.mt && (H.mt = N.mt), N.ct && (H.ct = N.ct)), H;
    }
    function mi(b, k) {
      W(b);
      var I = se.find(b, k);
      if (I) {
        for (var N = 0; N < b.FileIndex.length; ++N) if (b.FileIndex[N] == I)
          return b.FileIndex.splice(N, 1), b.FullPaths.splice(N, 1), !0;
      }
      return !1;
    }
    function wi(b, k, I) {
      W(b);
      var N = se.find(b, k);
      if (N) {
        for (var L = 0; L < b.FileIndex.length; ++L) if (b.FileIndex[L] == N)
          return b.FileIndex[L].name = M(I), b.FullPaths[L] = I, !0;
      }
      return !1;
    }
    function _i(b) {
      $(b, !0);
    }
    return G.find = he, G.read = z, G.parse = Te, G.write = Fe, G.writeFile = Xe, G.utils = {
      cfb_new: gi,
      cfb_add: Hr,
      cfb_del: mi,
      cfb_mov: wi,
      cfb_gc: _i,
      ReadShift: O,
      CheckField: T,
      prep_blob: j,
      bconcat: A,
      use_zlib: Bt,
      _deflateRaw: bn,
      _inflateRaw: Nn,
      consts: Je
    }, G;
  }();
  typeof jt < "u" && typeof DO_NOT_EXPORT_CFB > "u" && (t.exports = se);
})(wa);
var lt = wa.exports;
/*! pako 2.1.0 https://github.com/nodeca/pako @license (MIT AND Zlib) */
const Yi = 4, Cn = 0, Pn = 1, qi = 2;
function Ft(t) {
  let r = t.length;
  for (; --r >= 0; )
    t[r] = 0;
}
const Qi = 0, _a = 1, es = 2, ts = 3, rs = 258, fn = 29, ir = 256, Yt = ir + 1 + fn, Rt = 30, dn = 19, ya = 2 * Yt + 1, mt = 15, Mr = 16, ns = 7, pn = 256, Aa = 16, Sa = 17, ba = 18, en = (
  /* extra bits for each length code */
  new Uint8Array([0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0])
), br = (
  /* extra bits for each distance code */
  new Uint8Array([0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13])
), as = (
  /* extra bits for each bit length code */
  new Uint8Array([0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 3, 7])
), xa = new Uint8Array([16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15]), is = 512, st = new Array((Yt + 2) * 2);
Ft(st);
const Zt = new Array(Rt * 2);
Ft(Zt);
const qt = new Array(is);
Ft(qt);
const Qt = new Array(rs - ts + 1);
Ft(Qt);
const un = new Array(fn);
Ft(un);
const xr = new Array(Rt);
Ft(xr);
function $r(t, r, e, n, i) {
  this.static_tree = t, this.extra_bits = r, this.extra_base = e, this.elems = n, this.max_length = i, this.has_stree = t && t.length;
}
let Ta, Ea, Ia;
function Ur(t, r) {
  this.dyn_tree = t, this.max_code = 0, this.stat_desc = r;
}
const va = (t) => t < 256 ? qt[t] : qt[256 + (t >>> 7)], er = (t, r) => {
  t.pending_buf[t.pending++] = r & 255, t.pending_buf[t.pending++] = r >>> 8 & 255;
}, We = (t, r, e) => {
  t.bi_valid > Mr - e ? (t.bi_buf |= r << t.bi_valid & 65535, er(t, t.bi_buf), t.bi_buf = r >> Mr - t.bi_valid, t.bi_valid += e - Mr) : (t.bi_buf |= r << t.bi_valid & 65535, t.bi_valid += e);
}, tt = (t, r, e) => {
  We(
    t,
    e[r * 2],
    e[r * 2 + 1]
    /*.Len*/
  );
}, Na = (t, r) => {
  let e = 0;
  do
    e |= t & 1, t >>>= 1, e <<= 1;
  while (--r > 0);
  return e >>> 1;
}, ss = (t) => {
  t.bi_valid === 16 ? (er(t, t.bi_buf), t.bi_buf = 0, t.bi_valid = 0) : t.bi_valid >= 8 && (t.pending_buf[t.pending++] = t.bi_buf & 255, t.bi_buf >>= 8, t.bi_valid -= 8);
}, os = (t, r) => {
  const e = r.dyn_tree, n = r.max_code, i = r.stat_desc.static_tree, a = r.stat_desc.has_stree, s = r.stat_desc.extra_bits, o = r.stat_desc.extra_base, h = r.stat_desc.max_length;
  let l, f, d, u, c, g, p = 0;
  for (u = 0; u <= mt; u++)
    t.bl_count[u] = 0;
  for (e[t.heap[t.heap_max] * 2 + 1] = 0, l = t.heap_max + 1; l < ya; l++)
    f = t.heap[l], u = e[e[f * 2 + 1] * 2 + 1] + 1, u > h && (u = h, p++), e[f * 2 + 1] = u, !(f > n) && (t.bl_count[u]++, c = 0, f >= o && (c = s[f - o]), g = e[f * 2], t.opt_len += g * (u + c), a && (t.static_len += g * (i[f * 2 + 1] + c)));
  if (p !== 0) {
    do {
      for (u = h - 1; t.bl_count[u] === 0; )
        u--;
      t.bl_count[u]--, t.bl_count[u + 1] += 2, t.bl_count[h]--, p -= 2;
    } while (p > 0);
    for (u = h; u !== 0; u--)
      for (f = t.bl_count[u]; f !== 0; )
        d = t.heap[--l], !(d > n) && (e[d * 2 + 1] !== u && (t.opt_len += (u - e[d * 2 + 1]) * e[d * 2], e[d * 2 + 1] = u), f--);
  }
}, Ra = (t, r, e) => {
  const n = new Array(mt + 1);
  let i = 0, a, s;
  for (a = 1; a <= mt; a++)
    i = i + e[a - 1] << 1, n[a] = i;
  for (s = 0; s <= r; s++) {
    let o = t[s * 2 + 1];
    o !== 0 && (t[s * 2] = Na(n[o]++, o));
  }
}, ls = () => {
  let t, r, e, n, i;
  const a = new Array(mt + 1);
  for (e = 0, n = 0; n < fn - 1; n++)
    for (un[n] = e, t = 0; t < 1 << en[n]; t++)
      Qt[e++] = n;
  for (Qt[e - 1] = n, i = 0, n = 0; n < 16; n++)
    for (xr[n] = i, t = 0; t < 1 << br[n]; t++)
      qt[i++] = n;
  for (i >>= 7; n < Rt; n++)
    for (xr[n] = i << 7, t = 0; t < 1 << br[n] - 7; t++)
      qt[256 + i++] = n;
  for (r = 0; r <= mt; r++)
    a[r] = 0;
  for (t = 0; t <= 143; )
    st[t * 2 + 1] = 8, t++, a[8]++;
  for (; t <= 255; )
    st[t * 2 + 1] = 9, t++, a[9]++;
  for (; t <= 279; )
    st[t * 2 + 1] = 7, t++, a[7]++;
  for (; t <= 287; )
    st[t * 2 + 1] = 8, t++, a[8]++;
  for (Ra(st, Yt + 1, a), t = 0; t < Rt; t++)
    Zt[t * 2 + 1] = 5, Zt[t * 2] = Na(t, 5);
  Ta = new $r(st, en, ir + 1, Yt, mt), Ea = new $r(Zt, br, 0, Rt, mt), Ia = new $r(new Array(0), as, 0, dn, ns);
}, Ca = (t) => {
  let r;
  for (r = 0; r < Yt; r++)
    t.dyn_ltree[r * 2] = 0;
  for (r = 0; r < Rt; r++)
    t.dyn_dtree[r * 2] = 0;
  for (r = 0; r < dn; r++)
    t.bl_tree[r * 2] = 0;
  t.dyn_ltree[pn * 2] = 1, t.opt_len = t.static_len = 0, t.sym_next = t.matches = 0;
}, Pa = (t) => {
  t.bi_valid > 8 ? er(t, t.bi_buf) : t.bi_valid > 0 && (t.pending_buf[t.pending++] = t.bi_buf), t.bi_buf = 0, t.bi_valid = 0;
}, Fn = (t, r, e, n) => {
  const i = r * 2, a = e * 2;
  return t[i] < t[a] || t[i] === t[a] && n[r] <= n[e];
}, zr = (t, r, e) => {
  const n = t.heap[e];
  let i = e << 1;
  for (; i <= t.heap_len && (i < t.heap_len && Fn(r, t.heap[i + 1], t.heap[i], t.depth) && i++, !Fn(r, n, t.heap[i], t.depth)); )
    t.heap[e] = t.heap[i], e = i, i <<= 1;
  t.heap[e] = n;
}, kn = (t, r, e) => {
  let n, i, a = 0, s, o;
  if (t.sym_next !== 0)
    do
      n = t.pending_buf[t.sym_buf + a++] & 255, n += (t.pending_buf[t.sym_buf + a++] & 255) << 8, i = t.pending_buf[t.sym_buf + a++], n === 0 ? tt(t, i, r) : (s = Qt[i], tt(t, s + ir + 1, r), o = en[s], o !== 0 && (i -= un[s], We(t, i, o)), n--, s = va(n), tt(t, s, e), o = br[s], o !== 0 && (n -= xr[s], We(t, n, o)));
    while (a < t.sym_next);
  tt(t, pn, r);
}, tn = (t, r) => {
  const e = r.dyn_tree, n = r.stat_desc.static_tree, i = r.stat_desc.has_stree, a = r.stat_desc.elems;
  let s, o, h = -1, l;
  for (t.heap_len = 0, t.heap_max = ya, s = 0; s < a; s++)
    e[s * 2] !== 0 ? (t.heap[++t.heap_len] = h = s, t.depth[s] = 0) : e[s * 2 + 1] = 0;
  for (; t.heap_len < 2; )
    l = t.heap[++t.heap_len] = h < 2 ? ++h : 0, e[l * 2] = 1, t.depth[l] = 0, t.opt_len--, i && (t.static_len -= n[l * 2 + 1]);
  for (r.max_code = h, s = t.heap_len >> 1; s >= 1; s--)
    zr(t, e, s);
  l = a;
  do
    s = t.heap[
      1
      /*SMALLEST*/
    ], t.heap[
      1
      /*SMALLEST*/
    ] = t.heap[t.heap_len--], zr(
      t,
      e,
      1
      /*SMALLEST*/
    ), o = t.heap[
      1
      /*SMALLEST*/
    ], t.heap[--t.heap_max] = s, t.heap[--t.heap_max] = o, e[l * 2] = e[s * 2] + e[o * 2], t.depth[l] = (t.depth[s] >= t.depth[o] ? t.depth[s] : t.depth[o]) + 1, e[s * 2 + 1] = e[o * 2 + 1] = l, t.heap[
      1
      /*SMALLEST*/
    ] = l++, zr(
      t,
      e,
      1
      /*SMALLEST*/
    );
  while (t.heap_len >= 2);
  t.heap[--t.heap_max] = t.heap[
    1
    /*SMALLEST*/
  ], os(t, r), Ra(e, h, t.bl_count);
}, Ln = (t, r, e) => {
  let n, i = -1, a, s = r[0 * 2 + 1], o = 0, h = 7, l = 4;
  for (s === 0 && (h = 138, l = 3), r[(e + 1) * 2 + 1] = 65535, n = 0; n <= e; n++)
    a = s, s = r[(n + 1) * 2 + 1], !(++o < h && a === s) && (o < l ? t.bl_tree[a * 2] += o : a !== 0 ? (a !== i && t.bl_tree[a * 2]++, t.bl_tree[Aa * 2]++) : o <= 10 ? t.bl_tree[Sa * 2]++ : t.bl_tree[ba * 2]++, o = 0, i = a, s === 0 ? (h = 138, l = 3) : a === s ? (h = 6, l = 3) : (h = 7, l = 4));
}, Bn = (t, r, e) => {
  let n, i = -1, a, s = r[0 * 2 + 1], o = 0, h = 7, l = 4;
  for (s === 0 && (h = 138, l = 3), n = 0; n <= e; n++)
    if (a = s, s = r[(n + 1) * 2 + 1], !(++o < h && a === s)) {
      if (o < l)
        do
          tt(t, a, t.bl_tree);
        while (--o !== 0);
      else a !== 0 ? (a !== i && (tt(t, a, t.bl_tree), o--), tt(t, Aa, t.bl_tree), We(t, o - 3, 2)) : o <= 10 ? (tt(t, Sa, t.bl_tree), We(t, o - 3, 3)) : (tt(t, ba, t.bl_tree), We(t, o - 11, 7));
      o = 0, i = a, s === 0 ? (h = 138, l = 3) : a === s ? (h = 6, l = 3) : (h = 7, l = 4);
    }
}, hs = (t) => {
  let r;
  for (Ln(t, t.dyn_ltree, t.l_desc.max_code), Ln(t, t.dyn_dtree, t.d_desc.max_code), tn(t, t.bl_desc), r = dn - 1; r >= 3 && t.bl_tree[xa[r] * 2 + 1] === 0; r--)
    ;
  return t.opt_len += 3 * (r + 1) + 5 + 5 + 4, r;
}, cs = (t, r, e, n) => {
  let i;
  for (We(t, r - 257, 5), We(t, e - 1, 5), We(t, n - 4, 4), i = 0; i < n; i++)
    We(t, t.bl_tree[xa[i] * 2 + 1], 3);
  Bn(t, t.dyn_ltree, r - 1), Bn(t, t.dyn_dtree, e - 1);
}, fs = (t) => {
  let r = 4093624447, e;
  for (e = 0; e <= 31; e++, r >>>= 1)
    if (r & 1 && t.dyn_ltree[e * 2] !== 0)
      return Cn;
  if (t.dyn_ltree[9 * 2] !== 0 || t.dyn_ltree[10 * 2] !== 0 || t.dyn_ltree[13 * 2] !== 0)
    return Pn;
  for (e = 32; e < ir; e++)
    if (t.dyn_ltree[e * 2] !== 0)
      return Pn;
  return Cn;
};
let On = !1;
const ds = (t) => {
  On || (ls(), On = !0), t.l_desc = new Ur(t.dyn_ltree, Ta), t.d_desc = new Ur(t.dyn_dtree, Ea), t.bl_desc = new Ur(t.bl_tree, Ia), t.bi_buf = 0, t.bi_valid = 0, Ca(t);
}, Fa = (t, r, e, n) => {
  We(t, (Qi << 1) + (n ? 1 : 0), 3), Pa(t), er(t, e), er(t, ~e), e && t.pending_buf.set(t.window.subarray(r, r + e), t.pending), t.pending += e;
}, ps = (t) => {
  We(t, _a << 1, 3), tt(t, pn, st), ss(t);
}, us = (t, r, e, n) => {
  let i, a, s = 0;
  t.level > 0 ? (t.strm.data_type === qi && (t.strm.data_type = fs(t)), tn(t, t.l_desc), tn(t, t.d_desc), s = hs(t), i = t.opt_len + 3 + 7 >>> 3, a = t.static_len + 3 + 7 >>> 3, a <= i && (i = a)) : i = a = e + 5, e + 4 <= i && r !== -1 ? Fa(t, r, e, n) : t.strategy === Yi || a === i ? (We(t, (_a << 1) + (n ? 1 : 0), 3), kn(t, st, Zt)) : (We(t, (es << 1) + (n ? 1 : 0), 3), cs(t, t.l_desc.max_code + 1, t.d_desc.max_code + 1, s + 1), kn(t, t.dyn_ltree, t.dyn_dtree)), Ca(t), n && Pa(t);
}, gs = (t, r, e) => (t.pending_buf[t.sym_buf + t.sym_next++] = r, t.pending_buf[t.sym_buf + t.sym_next++] = r >> 8, t.pending_buf[t.sym_buf + t.sym_next++] = e, r === 0 ? t.dyn_ltree[e * 2]++ : (t.matches++, r--, t.dyn_ltree[(Qt[e] + ir + 1) * 2]++, t.dyn_dtree[va(r) * 2]++), t.sym_next === t.sym_end);
var ms = ds, ws = Fa, _s = us, ys = gs, As = ps, Ss = {
  _tr_init: ms,
  _tr_stored_block: ws,
  _tr_flush_block: _s,
  _tr_tally: ys,
  _tr_align: As
};
const bs = (t, r, e, n) => {
  let i = t & 65535 | 0, a = t >>> 16 & 65535 | 0, s = 0;
  for (; e !== 0; ) {
    s = e > 2e3 ? 2e3 : e, e -= s;
    do
      i = i + r[n++] | 0, a = a + i | 0;
    while (--s);
    i %= 65521, a %= 65521;
  }
  return i | a << 16 | 0;
};
var tr = bs;
const xs = () => {
  let t, r = [];
  for (var e = 0; e < 256; e++) {
    t = e;
    for (var n = 0; n < 8; n++)
      t = t & 1 ? 3988292384 ^ t >>> 1 : t >>> 1;
    r[e] = t;
  }
  return r;
}, Ts = new Uint32Array(xs()), Es = (t, r, e, n) => {
  const i = Ts, a = n + e;
  t ^= -1;
  for (let s = n; s < a; s++)
    t = t >>> 8 ^ i[(t ^ r[s]) & 255];
  return t ^ -1;
};
var Le = Es, At = {
  2: "need dictionary",
  /* Z_NEED_DICT       2  */
  1: "stream end",
  /* Z_STREAM_END      1  */
  0: "",
  /* Z_OK              0  */
  "-1": "file error",
  /* Z_ERRNO         (-1) */
  "-2": "stream error",
  /* Z_STREAM_ERROR  (-2) */
  "-3": "data error",
  /* Z_DATA_ERROR    (-3) */
  "-4": "insufficient memory",
  /* Z_MEM_ERROR     (-4) */
  "-5": "buffer error",
  /* Z_BUF_ERROR     (-5) */
  "-6": "incompatible version"
  /* Z_VERSION_ERROR (-6) */
}, sr = {
  /* Allowed flush values; see deflate() and inflate() below for details */
  Z_NO_FLUSH: 0,
  Z_PARTIAL_FLUSH: 1,
  Z_SYNC_FLUSH: 2,
  Z_FULL_FLUSH: 3,
  Z_FINISH: 4,
  Z_BLOCK: 5,
  Z_TREES: 6,
  /* Return codes for the compression/decompression functions. Negative values
  * are errors, positive values are used for special but normal events.
  */
  Z_OK: 0,
  Z_STREAM_END: 1,
  Z_NEED_DICT: 2,
  Z_ERRNO: -1,
  Z_STREAM_ERROR: -2,
  Z_DATA_ERROR: -3,
  Z_MEM_ERROR: -4,
  Z_BUF_ERROR: -5,
  //Z_VERSION_ERROR: -6,
  /* compression levels */
  Z_NO_COMPRESSION: 0,
  Z_BEST_SPEED: 1,
  Z_BEST_COMPRESSION: 9,
  Z_DEFAULT_COMPRESSION: -1,
  Z_FILTERED: 1,
  Z_HUFFMAN_ONLY: 2,
  Z_RLE: 3,
  Z_FIXED: 4,
  Z_DEFAULT_STRATEGY: 0,
  /* Possible values of the data_type field (though see inflate()) */
  Z_BINARY: 0,
  Z_TEXT: 1,
  //Z_ASCII:                1, // = Z_TEXT (deprecated)
  Z_UNKNOWN: 2,
  /* The deflate compression method */
  Z_DEFLATED: 8
  //Z_NULL:                 null // Use -1 or null inline, depending on var type
};
const { _tr_init: Is, _tr_stored_block: rn, _tr_flush_block: vs, _tr_tally: ft, _tr_align: Ns } = Ss, {
  Z_NO_FLUSH: dt,
  Z_PARTIAL_FLUSH: Rs,
  Z_FULL_FLUSH: Cs,
  Z_FINISH: Ze,
  Z_BLOCK: Hn,
  Z_OK: Be,
  Z_STREAM_END: Dn,
  Z_STREAM_ERROR: rt,
  Z_DATA_ERROR: Ps,
  Z_BUF_ERROR: Wr,
  Z_DEFAULT_COMPRESSION: Fs,
  Z_FILTERED: ks,
  Z_HUFFMAN_ONLY: wr,
  Z_RLE: Ls,
  Z_FIXED: Bs,
  Z_DEFAULT_STRATEGY: Os,
  Z_UNKNOWN: Hs,
  Z_DEFLATED: Nr
} = sr, Ds = 9, Ms = 15, $s = 8, Us = 29, zs = 256, nn = zs + 1 + Us, Ws = 30, js = 19, Gs = 2 * nn + 1, Xs = 15, xe = 3, ct = 258, nt = ct + xe + 1, Zs = 32, Ct = 42, gn = 57, an = 69, sn = 73, on = 91, ln = 103, wt = 113, Gt = 666, $e = 1, kt = 2, St = 3, Lt = 4, Ks = 3, _t = (t, r) => (t.msg = At[r], r), Mn = (t) => t * 2 - (t > 4 ? 9 : 0), ht = (t) => {
  let r = t.length;
  for (; --r >= 0; )
    t[r] = 0;
}, Js = (t) => {
  let r, e, n, i = t.w_size;
  r = t.hash_size, n = r;
  do
    e = t.head[--n], t.head[n] = e >= i ? e - i : 0;
  while (--r);
  r = i, n = r;
  do
    e = t.prev[--n], t.prev[n] = e >= i ? e - i : 0;
  while (--r);
};
let Vs = (t, r, e) => (r << t.hash_shift ^ e) & t.hash_mask, pt = Vs;
const je = (t) => {
  const r = t.state;
  let e = r.pending;
  e > t.avail_out && (e = t.avail_out), e !== 0 && (t.output.set(r.pending_buf.subarray(r.pending_out, r.pending_out + e), t.next_out), t.next_out += e, r.pending_out += e, t.total_out += e, t.avail_out -= e, r.pending -= e, r.pending === 0 && (r.pending_out = 0));
}, Ge = (t, r) => {
  vs(t, t.block_start >= 0 ? t.block_start : -1, t.strstart - t.block_start, r), t.block_start = t.strstart, je(t.strm);
}, Ee = (t, r) => {
  t.pending_buf[t.pending++] = r;
}, Ut = (t, r) => {
  t.pending_buf[t.pending++] = r >>> 8 & 255, t.pending_buf[t.pending++] = r & 255;
}, hn = (t, r, e, n) => {
  let i = t.avail_in;
  return i > n && (i = n), i === 0 ? 0 : (t.avail_in -= i, r.set(t.input.subarray(t.next_in, t.next_in + i), e), t.state.wrap === 1 ? t.adler = tr(t.adler, r, i, e) : t.state.wrap === 2 && (t.adler = Le(t.adler, r, i, e)), t.next_in += i, t.total_in += i, i);
}, ka = (t, r) => {
  let e = t.max_chain_length, n = t.strstart, i, a, s = t.prev_length, o = t.nice_match;
  const h = t.strstart > t.w_size - nt ? t.strstart - (t.w_size - nt) : 0, l = t.window, f = t.w_mask, d = t.prev, u = t.strstart + ct;
  let c = l[n + s - 1], g = l[n + s];
  t.prev_length >= t.good_match && (e >>= 2), o > t.lookahead && (o = t.lookahead);
  do
    if (i = r, !(l[i + s] !== g || l[i + s - 1] !== c || l[i] !== l[n] || l[++i] !== l[n + 1])) {
      n += 2, i++;
      do
        ;
      while (l[++n] === l[++i] && l[++n] === l[++i] && l[++n] === l[++i] && l[++n] === l[++i] && l[++n] === l[++i] && l[++n] === l[++i] && l[++n] === l[++i] && l[++n] === l[++i] && n < u);
      if (a = ct - (u - n), n = u - ct, a > s) {
        if (t.match_start = r, s = a, a >= o)
          break;
        c = l[n + s - 1], g = l[n + s];
      }
    }
  while ((r = d[r & f]) > h && --e !== 0);
  return s <= t.lookahead ? s : t.lookahead;
}, Pt = (t) => {
  const r = t.w_size;
  let e, n, i;
  do {
    if (n = t.window_size - t.lookahead - t.strstart, t.strstart >= r + (r - nt) && (t.window.set(t.window.subarray(r, r + r - n), 0), t.match_start -= r, t.strstart -= r, t.block_start -= r, t.insert > t.strstart && (t.insert = t.strstart), Js(t), n += r), t.strm.avail_in === 0)
      break;
    if (e = hn(t.strm, t.window, t.strstart + t.lookahead, n), t.lookahead += e, t.lookahead + t.insert >= xe)
      for (i = t.strstart - t.insert, t.ins_h = t.window[i], t.ins_h = pt(t, t.ins_h, t.window[i + 1]); t.insert && (t.ins_h = pt(t, t.ins_h, t.window[i + xe - 1]), t.prev[i & t.w_mask] = t.head[t.ins_h], t.head[t.ins_h] = i, i++, t.insert--, !(t.lookahead + t.insert < xe)); )
        ;
  } while (t.lookahead < nt && t.strm.avail_in !== 0);
}, La = (t, r) => {
  let e = t.pending_buf_size - 5 > t.w_size ? t.w_size : t.pending_buf_size - 5, n, i, a, s = 0, o = t.strm.avail_in;
  do {
    if (n = 65535, a = t.bi_valid + 42 >> 3, t.strm.avail_out < a || (a = t.strm.avail_out - a, i = t.strstart - t.block_start, n > i + t.strm.avail_in && (n = i + t.strm.avail_in), n > a && (n = a), n < e && (n === 0 && r !== Ze || r === dt || n !== i + t.strm.avail_in)))
      break;
    s = r === Ze && n === i + t.strm.avail_in ? 1 : 0, rn(t, 0, 0, s), t.pending_buf[t.pending - 4] = n, t.pending_buf[t.pending - 3] = n >> 8, t.pending_buf[t.pending - 2] = ~n, t.pending_buf[t.pending - 1] = ~n >> 8, je(t.strm), i && (i > n && (i = n), t.strm.output.set(t.window.subarray(t.block_start, t.block_start + i), t.strm.next_out), t.strm.next_out += i, t.strm.avail_out -= i, t.strm.total_out += i, t.block_start += i, n -= i), n && (hn(t.strm, t.strm.output, t.strm.next_out, n), t.strm.next_out += n, t.strm.avail_out -= n, t.strm.total_out += n);
  } while (s === 0);
  return o -= t.strm.avail_in, o && (o >= t.w_size ? (t.matches = 2, t.window.set(t.strm.input.subarray(t.strm.next_in - t.w_size, t.strm.next_in), 0), t.strstart = t.w_size, t.insert = t.strstart) : (t.window_size - t.strstart <= o && (t.strstart -= t.w_size, t.window.set(t.window.subarray(t.w_size, t.w_size + t.strstart), 0), t.matches < 2 && t.matches++, t.insert > t.strstart && (t.insert = t.strstart)), t.window.set(t.strm.input.subarray(t.strm.next_in - o, t.strm.next_in), t.strstart), t.strstart += o, t.insert += o > t.w_size - t.insert ? t.w_size - t.insert : o), t.block_start = t.strstart), t.high_water < t.strstart && (t.high_water = t.strstart), s ? Lt : r !== dt && r !== Ze && t.strm.avail_in === 0 && t.strstart === t.block_start ? kt : (a = t.window_size - t.strstart, t.strm.avail_in > a && t.block_start >= t.w_size && (t.block_start -= t.w_size, t.strstart -= t.w_size, t.window.set(t.window.subarray(t.w_size, t.w_size + t.strstart), 0), t.matches < 2 && t.matches++, a += t.w_size, t.insert > t.strstart && (t.insert = t.strstart)), a > t.strm.avail_in && (a = t.strm.avail_in), a && (hn(t.strm, t.window, t.strstart, a), t.strstart += a, t.insert += a > t.w_size - t.insert ? t.w_size - t.insert : a), t.high_water < t.strstart && (t.high_water = t.strstart), a = t.bi_valid + 42 >> 3, a = t.pending_buf_size - a > 65535 ? 65535 : t.pending_buf_size - a, e = a > t.w_size ? t.w_size : a, i = t.strstart - t.block_start, (i >= e || (i || r === Ze) && r !== dt && t.strm.avail_in === 0 && i <= a) && (n = i > a ? a : i, s = r === Ze && t.strm.avail_in === 0 && n === i ? 1 : 0, rn(t, t.block_start, n, s), t.block_start += n, je(t.strm)), s ? St : $e);
}, jr = (t, r) => {
  let e, n;
  for (; ; ) {
    if (t.lookahead < nt) {
      if (Pt(t), t.lookahead < nt && r === dt)
        return $e;
      if (t.lookahead === 0)
        break;
    }
    if (e = 0, t.lookahead >= xe && (t.ins_h = pt(t, t.ins_h, t.window[t.strstart + xe - 1]), e = t.prev[t.strstart & t.w_mask] = t.head[t.ins_h], t.head[t.ins_h] = t.strstart), e !== 0 && t.strstart - e <= t.w_size - nt && (t.match_length = ka(t, e)), t.match_length >= xe)
      if (n = ft(t, t.strstart - t.match_start, t.match_length - xe), t.lookahead -= t.match_length, t.match_length <= t.max_lazy_match && t.lookahead >= xe) {
        t.match_length--;
        do
          t.strstart++, t.ins_h = pt(t, t.ins_h, t.window[t.strstart + xe - 1]), e = t.prev[t.strstart & t.w_mask] = t.head[t.ins_h], t.head[t.ins_h] = t.strstart;
        while (--t.match_length !== 0);
        t.strstart++;
      } else
        t.strstart += t.match_length, t.match_length = 0, t.ins_h = t.window[t.strstart], t.ins_h = pt(t, t.ins_h, t.window[t.strstart + 1]);
    else
      n = ft(t, 0, t.window[t.strstart]), t.lookahead--, t.strstart++;
    if (n && (Ge(t, !1), t.strm.avail_out === 0))
      return $e;
  }
  return t.insert = t.strstart < xe - 1 ? t.strstart : xe - 1, r === Ze ? (Ge(t, !0), t.strm.avail_out === 0 ? St : Lt) : t.sym_next && (Ge(t, !1), t.strm.avail_out === 0) ? $e : kt;
}, vt = (t, r) => {
  let e, n, i;
  for (; ; ) {
    if (t.lookahead < nt) {
      if (Pt(t), t.lookahead < nt && r === dt)
        return $e;
      if (t.lookahead === 0)
        break;
    }
    if (e = 0, t.lookahead >= xe && (t.ins_h = pt(t, t.ins_h, t.window[t.strstart + xe - 1]), e = t.prev[t.strstart & t.w_mask] = t.head[t.ins_h], t.head[t.ins_h] = t.strstart), t.prev_length = t.match_length, t.prev_match = t.match_start, t.match_length = xe - 1, e !== 0 && t.prev_length < t.max_lazy_match && t.strstart - e <= t.w_size - nt && (t.match_length = ka(t, e), t.match_length <= 5 && (t.strategy === ks || t.match_length === xe && t.strstart - t.match_start > 4096) && (t.match_length = xe - 1)), t.prev_length >= xe && t.match_length <= t.prev_length) {
      i = t.strstart + t.lookahead - xe, n = ft(t, t.strstart - 1 - t.prev_match, t.prev_length - xe), t.lookahead -= t.prev_length - 1, t.prev_length -= 2;
      do
        ++t.strstart <= i && (t.ins_h = pt(t, t.ins_h, t.window[t.strstart + xe - 1]), e = t.prev[t.strstart & t.w_mask] = t.head[t.ins_h], t.head[t.ins_h] = t.strstart);
      while (--t.prev_length !== 0);
      if (t.match_available = 0, t.match_length = xe - 1, t.strstart++, n && (Ge(t, !1), t.strm.avail_out === 0))
        return $e;
    } else if (t.match_available) {
      if (n = ft(t, 0, t.window[t.strstart - 1]), n && Ge(t, !1), t.strstart++, t.lookahead--, t.strm.avail_out === 0)
        return $e;
    } else
      t.match_available = 1, t.strstart++, t.lookahead--;
  }
  return t.match_available && (n = ft(t, 0, t.window[t.strstart - 1]), t.match_available = 0), t.insert = t.strstart < xe - 1 ? t.strstart : xe - 1, r === Ze ? (Ge(t, !0), t.strm.avail_out === 0 ? St : Lt) : t.sym_next && (Ge(t, !1), t.strm.avail_out === 0) ? $e : kt;
}, Ys = (t, r) => {
  let e, n, i, a;
  const s = t.window;
  for (; ; ) {
    if (t.lookahead <= ct) {
      if (Pt(t), t.lookahead <= ct && r === dt)
        return $e;
      if (t.lookahead === 0)
        break;
    }
    if (t.match_length = 0, t.lookahead >= xe && t.strstart > 0 && (i = t.strstart - 1, n = s[i], n === s[++i] && n === s[++i] && n === s[++i])) {
      a = t.strstart + ct;
      do
        ;
      while (n === s[++i] && n === s[++i] && n === s[++i] && n === s[++i] && n === s[++i] && n === s[++i] && n === s[++i] && n === s[++i] && i < a);
      t.match_length = ct - (a - i), t.match_length > t.lookahead && (t.match_length = t.lookahead);
    }
    if (t.match_length >= xe ? (e = ft(t, 1, t.match_length - xe), t.lookahead -= t.match_length, t.strstart += t.match_length, t.match_length = 0) : (e = ft(t, 0, t.window[t.strstart]), t.lookahead--, t.strstart++), e && (Ge(t, !1), t.strm.avail_out === 0))
      return $e;
  }
  return t.insert = 0, r === Ze ? (Ge(t, !0), t.strm.avail_out === 0 ? St : Lt) : t.sym_next && (Ge(t, !1), t.strm.avail_out === 0) ? $e : kt;
}, qs = (t, r) => {
  let e;
  for (; ; ) {
    if (t.lookahead === 0 && (Pt(t), t.lookahead === 0)) {
      if (r === dt)
        return $e;
      break;
    }
    if (t.match_length = 0, e = ft(t, 0, t.window[t.strstart]), t.lookahead--, t.strstart++, e && (Ge(t, !1), t.strm.avail_out === 0))
      return $e;
  }
  return t.insert = 0, r === Ze ? (Ge(t, !0), t.strm.avail_out === 0 ? St : Lt) : t.sym_next && (Ge(t, !1), t.strm.avail_out === 0) ? $e : kt;
};
function et(t, r, e, n, i) {
  this.good_length = t, this.max_lazy = r, this.nice_length = e, this.max_chain = n, this.func = i;
}
const Xt = [
  /*      good lazy nice chain */
  new et(0, 0, 0, 0, La),
  /* 0 store only */
  new et(4, 4, 8, 4, jr),
  /* 1 max speed, no lazy matches */
  new et(4, 5, 16, 8, jr),
  /* 2 */
  new et(4, 6, 32, 32, jr),
  /* 3 */
  new et(4, 4, 16, 16, vt),
  /* 4 lazy matches */
  new et(8, 16, 32, 32, vt),
  /* 5 */
  new et(8, 16, 128, 128, vt),
  /* 6 */
  new et(8, 32, 128, 256, vt),
  /* 7 */
  new et(32, 128, 258, 1024, vt),
  /* 8 */
  new et(32, 258, 258, 4096, vt)
  /* 9 max compression */
], Qs = (t) => {
  t.window_size = 2 * t.w_size, ht(t.head), t.max_lazy_match = Xt[t.level].max_lazy, t.good_match = Xt[t.level].good_length, t.nice_match = Xt[t.level].nice_length, t.max_chain_length = Xt[t.level].max_chain, t.strstart = 0, t.block_start = 0, t.lookahead = 0, t.insert = 0, t.match_length = t.prev_length = xe - 1, t.match_available = 0, t.ins_h = 0;
};
function eo() {
  this.strm = null, this.status = 0, this.pending_buf = null, this.pending_buf_size = 0, this.pending_out = 0, this.pending = 0, this.wrap = 0, this.gzhead = null, this.gzindex = 0, this.method = Nr, this.last_flush = -1, this.w_size = 0, this.w_bits = 0, this.w_mask = 0, this.window = null, this.window_size = 0, this.prev = null, this.head = null, this.ins_h = 0, this.hash_size = 0, this.hash_bits = 0, this.hash_mask = 0, this.hash_shift = 0, this.block_start = 0, this.match_length = 0, this.prev_match = 0, this.match_available = 0, this.strstart = 0, this.match_start = 0, this.lookahead = 0, this.prev_length = 0, this.max_chain_length = 0, this.max_lazy_match = 0, this.level = 0, this.strategy = 0, this.good_match = 0, this.nice_match = 0, this.dyn_ltree = new Uint16Array(Gs * 2), this.dyn_dtree = new Uint16Array((2 * Ws + 1) * 2), this.bl_tree = new Uint16Array((2 * js + 1) * 2), ht(this.dyn_ltree), ht(this.dyn_dtree), ht(this.bl_tree), this.l_desc = null, this.d_desc = null, this.bl_desc = null, this.bl_count = new Uint16Array(Xs + 1), this.heap = new Uint16Array(2 * nn + 1), ht(this.heap), this.heap_len = 0, this.heap_max = 0, this.depth = new Uint16Array(2 * nn + 1), ht(this.depth), this.sym_buf = 0, this.lit_bufsize = 0, this.sym_next = 0, this.sym_end = 0, this.opt_len = 0, this.static_len = 0, this.matches = 0, this.insert = 0, this.bi_buf = 0, this.bi_valid = 0;
}
const or = (t) => {
  if (!t)
    return 1;
  const r = t.state;
  return !r || r.strm !== t || r.status !== Ct && //#ifdef GZIP
  r.status !== gn && //#endif
  r.status !== an && r.status !== sn && r.status !== on && r.status !== ln && r.status !== wt && r.status !== Gt ? 1 : 0;
}, Ba = (t) => {
  if (or(t))
    return _t(t, rt);
  t.total_in = t.total_out = 0, t.data_type = Hs;
  const r = t.state;
  return r.pending = 0, r.pending_out = 0, r.wrap < 0 && (r.wrap = -r.wrap), r.status = //#ifdef GZIP
  r.wrap === 2 ? gn : (
    //#endif
    r.wrap ? Ct : wt
  ), t.adler = r.wrap === 2 ? 0 : 1, r.last_flush = -2, Is(r), Be;
}, Oa = (t) => {
  const r = Ba(t);
  return r === Be && Qs(t.state), r;
}, to = (t, r) => or(t) || t.state.wrap !== 2 ? rt : (t.state.gzhead = r, Be), Ha = (t, r, e, n, i, a) => {
  if (!t)
    return rt;
  let s = 1;
  if (r === Fs && (r = 6), n < 0 ? (s = 0, n = -n) : n > 15 && (s = 2, n -= 16), i < 1 || i > Ds || e !== Nr || n < 8 || n > 15 || r < 0 || r > 9 || a < 0 || a > Bs || n === 8 && s !== 1)
    return _t(t, rt);
  n === 8 && (n = 9);
  const o = new eo();
  return t.state = o, o.strm = t, o.status = Ct, o.wrap = s, o.gzhead = null, o.w_bits = n, o.w_size = 1 << o.w_bits, o.w_mask = o.w_size - 1, o.hash_bits = i + 7, o.hash_size = 1 << o.hash_bits, o.hash_mask = o.hash_size - 1, o.hash_shift = ~~((o.hash_bits + xe - 1) / xe), o.window = new Uint8Array(o.w_size * 2), o.head = new Uint16Array(o.hash_size), o.prev = new Uint16Array(o.w_size), o.lit_bufsize = 1 << i + 6, o.pending_buf_size = o.lit_bufsize * 4, o.pending_buf = new Uint8Array(o.pending_buf_size), o.sym_buf = o.lit_bufsize, o.sym_end = (o.lit_bufsize - 1) * 3, o.level = r, o.strategy = a, o.method = e, Oa(t);
}, ro = (t, r) => Ha(t, r, Nr, Ms, $s, Os), no = (t, r) => {
  if (or(t) || r > Hn || r < 0)
    return t ? _t(t, rt) : rt;
  const e = t.state;
  if (!t.output || t.avail_in !== 0 && !t.input || e.status === Gt && r !== Ze)
    return _t(t, t.avail_out === 0 ? Wr : rt);
  const n = e.last_flush;
  if (e.last_flush = r, e.pending !== 0) {
    if (je(t), t.avail_out === 0)
      return e.last_flush = -1, Be;
  } else if (t.avail_in === 0 && Mn(r) <= Mn(n) && r !== Ze)
    return _t(t, Wr);
  if (e.status === Gt && t.avail_in !== 0)
    return _t(t, Wr);
  if (e.status === Ct && e.wrap === 0 && (e.status = wt), e.status === Ct) {
    let i = Nr + (e.w_bits - 8 << 4) << 8, a = -1;
    if (e.strategy >= wr || e.level < 2 ? a = 0 : e.level < 6 ? a = 1 : e.level === 6 ? a = 2 : a = 3, i |= a << 6, e.strstart !== 0 && (i |= Zs), i += 31 - i % 31, Ut(e, i), e.strstart !== 0 && (Ut(e, t.adler >>> 16), Ut(e, t.adler & 65535)), t.adler = 1, e.status = wt, je(t), e.pending !== 0)
      return e.last_flush = -1, Be;
  }
  if (e.status === gn) {
    if (t.adler = 0, Ee(e, 31), Ee(e, 139), Ee(e, 8), e.gzhead)
      Ee(
        e,
        (e.gzhead.text ? 1 : 0) + (e.gzhead.hcrc ? 2 : 0) + (e.gzhead.extra ? 4 : 0) + (e.gzhead.name ? 8 : 0) + (e.gzhead.comment ? 16 : 0)
      ), Ee(e, e.gzhead.time & 255), Ee(e, e.gzhead.time >> 8 & 255), Ee(e, e.gzhead.time >> 16 & 255), Ee(e, e.gzhead.time >> 24 & 255), Ee(e, e.level === 9 ? 2 : e.strategy >= wr || e.level < 2 ? 4 : 0), Ee(e, e.gzhead.os & 255), e.gzhead.extra && e.gzhead.extra.length && (Ee(e, e.gzhead.extra.length & 255), Ee(e, e.gzhead.extra.length >> 8 & 255)), e.gzhead.hcrc && (t.adler = Le(t.adler, e.pending_buf, e.pending, 0)), e.gzindex = 0, e.status = an;
    else if (Ee(e, 0), Ee(e, 0), Ee(e, 0), Ee(e, 0), Ee(e, 0), Ee(e, e.level === 9 ? 2 : e.strategy >= wr || e.level < 2 ? 4 : 0), Ee(e, Ks), e.status = wt, je(t), e.pending !== 0)
      return e.last_flush = -1, Be;
  }
  if (e.status === an) {
    if (e.gzhead.extra) {
      let i = e.pending, a = (e.gzhead.extra.length & 65535) - e.gzindex;
      for (; e.pending + a > e.pending_buf_size; ) {
        let o = e.pending_buf_size - e.pending;
        if (e.pending_buf.set(e.gzhead.extra.subarray(e.gzindex, e.gzindex + o), e.pending), e.pending = e.pending_buf_size, e.gzhead.hcrc && e.pending > i && (t.adler = Le(t.adler, e.pending_buf, e.pending - i, i)), e.gzindex += o, je(t), e.pending !== 0)
          return e.last_flush = -1, Be;
        i = 0, a -= o;
      }
      let s = new Uint8Array(e.gzhead.extra);
      e.pending_buf.set(s.subarray(e.gzindex, e.gzindex + a), e.pending), e.pending += a, e.gzhead.hcrc && e.pending > i && (t.adler = Le(t.adler, e.pending_buf, e.pending - i, i)), e.gzindex = 0;
    }
    e.status = sn;
  }
  if (e.status === sn) {
    if (e.gzhead.name) {
      let i = e.pending, a;
      do {
        if (e.pending === e.pending_buf_size) {
          if (e.gzhead.hcrc && e.pending > i && (t.adler = Le(t.adler, e.pending_buf, e.pending - i, i)), je(t), e.pending !== 0)
            return e.last_flush = -1, Be;
          i = 0;
        }
        e.gzindex < e.gzhead.name.length ? a = e.gzhead.name.charCodeAt(e.gzindex++) & 255 : a = 0, Ee(e, a);
      } while (a !== 0);
      e.gzhead.hcrc && e.pending > i && (t.adler = Le(t.adler, e.pending_buf, e.pending - i, i)), e.gzindex = 0;
    }
    e.status = on;
  }
  if (e.status === on) {
    if (e.gzhead.comment) {
      let i = e.pending, a;
      do {
        if (e.pending === e.pending_buf_size) {
          if (e.gzhead.hcrc && e.pending > i && (t.adler = Le(t.adler, e.pending_buf, e.pending - i, i)), je(t), e.pending !== 0)
            return e.last_flush = -1, Be;
          i = 0;
        }
        e.gzindex < e.gzhead.comment.length ? a = e.gzhead.comment.charCodeAt(e.gzindex++) & 255 : a = 0, Ee(e, a);
      } while (a !== 0);
      e.gzhead.hcrc && e.pending > i && (t.adler = Le(t.adler, e.pending_buf, e.pending - i, i));
    }
    e.status = ln;
  }
  if (e.status === ln) {
    if (e.gzhead.hcrc) {
      if (e.pending + 2 > e.pending_buf_size && (je(t), e.pending !== 0))
        return e.last_flush = -1, Be;
      Ee(e, t.adler & 255), Ee(e, t.adler >> 8 & 255), t.adler = 0;
    }
    if (e.status = wt, je(t), e.pending !== 0)
      return e.last_flush = -1, Be;
  }
  if (t.avail_in !== 0 || e.lookahead !== 0 || r !== dt && e.status !== Gt) {
    let i = e.level === 0 ? La(e, r) : e.strategy === wr ? qs(e, r) : e.strategy === Ls ? Ys(e, r) : Xt[e.level].func(e, r);
    if ((i === St || i === Lt) && (e.status = Gt), i === $e || i === St)
      return t.avail_out === 0 && (e.last_flush = -1), Be;
    if (i === kt && (r === Rs ? Ns(e) : r !== Hn && (rn(e, 0, 0, !1), r === Cs && (ht(e.head), e.lookahead === 0 && (e.strstart = 0, e.block_start = 0, e.insert = 0))), je(t), t.avail_out === 0))
      return e.last_flush = -1, Be;
  }
  return r !== Ze ? Be : e.wrap <= 0 ? Dn : (e.wrap === 2 ? (Ee(e, t.adler & 255), Ee(e, t.adler >> 8 & 255), Ee(e, t.adler >> 16 & 255), Ee(e, t.adler >> 24 & 255), Ee(e, t.total_in & 255), Ee(e, t.total_in >> 8 & 255), Ee(e, t.total_in >> 16 & 255), Ee(e, t.total_in >> 24 & 255)) : (Ut(e, t.adler >>> 16), Ut(e, t.adler & 65535)), je(t), e.wrap > 0 && (e.wrap = -e.wrap), e.pending !== 0 ? Be : Dn);
}, ao = (t) => {
  if (or(t))
    return rt;
  const r = t.state.status;
  return t.state = null, r === wt ? _t(t, Ps) : Be;
}, io = (t, r) => {
  let e = r.length;
  if (or(t))
    return rt;
  const n = t.state, i = n.wrap;
  if (i === 2 || i === 1 && n.status !== Ct || n.lookahead)
    return rt;
  if (i === 1 && (t.adler = tr(t.adler, r, e, 0)), n.wrap = 0, e >= n.w_size) {
    i === 0 && (ht(n.head), n.strstart = 0, n.block_start = 0, n.insert = 0);
    let h = new Uint8Array(n.w_size);
    h.set(r.subarray(e - n.w_size, e), 0), r = h, e = n.w_size;
  }
  const a = t.avail_in, s = t.next_in, o = t.input;
  for (t.avail_in = e, t.next_in = 0, t.input = r, Pt(n); n.lookahead >= xe; ) {
    let h = n.strstart, l = n.lookahead - (xe - 1);
    do
      n.ins_h = pt(n, n.ins_h, n.window[h + xe - 1]), n.prev[h & n.w_mask] = n.head[n.ins_h], n.head[n.ins_h] = h, h++;
    while (--l);
    n.strstart = h, n.lookahead = xe - 1, Pt(n);
  }
  return n.strstart += n.lookahead, n.block_start = n.strstart, n.insert = n.lookahead, n.lookahead = 0, n.match_length = n.prev_length = xe - 1, n.match_available = 0, t.next_in = s, t.input = o, t.avail_in = a, n.wrap = i, Be;
};
var so = ro, oo = Ha, lo = Oa, ho = Ba, co = to, fo = no, po = ao, uo = io, go = "pako deflate (from Nodeca project)", Kt = {
  deflateInit: so,
  deflateInit2: oo,
  deflateReset: lo,
  deflateResetKeep: ho,
  deflateSetHeader: co,
  deflate: fo,
  deflateEnd: po,
  deflateSetDictionary: uo,
  deflateInfo: go
};
const mo = (t, r) => Object.prototype.hasOwnProperty.call(t, r);
var wo = function(t) {
  const r = Array.prototype.slice.call(arguments, 1);
  for (; r.length; ) {
    const e = r.shift();
    if (e) {
      if (typeof e != "object")
        throw new TypeError(e + "must be non-object");
      for (const n in e)
        mo(e, n) && (t[n] = e[n]);
    }
  }
  return t;
}, _o = (t) => {
  let r = 0;
  for (let n = 0, i = t.length; n < i; n++)
    r += t[n].length;
  const e = new Uint8Array(r);
  for (let n = 0, i = 0, a = t.length; n < a; n++) {
    let s = t[n];
    e.set(s, i), i += s.length;
  }
  return e;
}, Rr = {
  assign: wo,
  flattenChunks: _o
};
let Da = !0;
try {
  String.fromCharCode.apply(null, new Uint8Array(1));
} catch {
  Da = !1;
}
const rr = new Uint8Array(256);
for (let t = 0; t < 256; t++)
  rr[t] = t >= 252 ? 6 : t >= 248 ? 5 : t >= 240 ? 4 : t >= 224 ? 3 : t >= 192 ? 2 : 1;
rr[254] = rr[254] = 1;
var yo = (t) => {
  if (typeof TextEncoder == "function" && TextEncoder.prototype.encode)
    return new TextEncoder().encode(t);
  let r, e, n, i, a, s = t.length, o = 0;
  for (i = 0; i < s; i++)
    e = t.charCodeAt(i), (e & 64512) === 55296 && i + 1 < s && (n = t.charCodeAt(i + 1), (n & 64512) === 56320 && (e = 65536 + (e - 55296 << 10) + (n - 56320), i++)), o += e < 128 ? 1 : e < 2048 ? 2 : e < 65536 ? 3 : 4;
  for (r = new Uint8Array(o), a = 0, i = 0; a < o; i++)
    e = t.charCodeAt(i), (e & 64512) === 55296 && i + 1 < s && (n = t.charCodeAt(i + 1), (n & 64512) === 56320 && (e = 65536 + (e - 55296 << 10) + (n - 56320), i++)), e < 128 ? r[a++] = e : e < 2048 ? (r[a++] = 192 | e >>> 6, r[a++] = 128 | e & 63) : e < 65536 ? (r[a++] = 224 | e >>> 12, r[a++] = 128 | e >>> 6 & 63, r[a++] = 128 | e & 63) : (r[a++] = 240 | e >>> 18, r[a++] = 128 | e >>> 12 & 63, r[a++] = 128 | e >>> 6 & 63, r[a++] = 128 | e & 63);
  return r;
};
const Ao = (t, r) => {
  if (r < 65534 && t.subarray && Da)
    return String.fromCharCode.apply(null, t.length === r ? t : t.subarray(0, r));
  let e = "";
  for (let n = 0; n < r; n++)
    e += String.fromCharCode(t[n]);
  return e;
};
var So = (t, r) => {
  const e = r || t.length;
  if (typeof TextDecoder == "function" && TextDecoder.prototype.decode)
    return new TextDecoder().decode(t.subarray(0, r));
  let n, i;
  const a = new Array(e * 2);
  for (i = 0, n = 0; n < e; ) {
    let s = t[n++];
    if (s < 128) {
      a[i++] = s;
      continue;
    }
    let o = rr[s];
    if (o > 4) {
      a[i++] = 65533, n += o - 1;
      continue;
    }
    for (s &= o === 2 ? 31 : o === 3 ? 15 : 7; o > 1 && n < e; )
      s = s << 6 | t[n++] & 63, o--;
    if (o > 1) {
      a[i++] = 65533;
      continue;
    }
    s < 65536 ? a[i++] = s : (s -= 65536, a[i++] = 55296 | s >> 10 & 1023, a[i++] = 56320 | s & 1023);
  }
  return Ao(a, i);
}, bo = (t, r) => {
  r = r || t.length, r > t.length && (r = t.length);
  let e = r - 1;
  for (; e >= 0 && (t[e] & 192) === 128; )
    e--;
  return e < 0 || e === 0 ? r : e + rr[t[e]] > r ? e : r;
}, nr = {
  string2buf: yo,
  buf2string: So,
  utf8border: bo
};
function xo() {
  this.input = null, this.next_in = 0, this.avail_in = 0, this.total_in = 0, this.output = null, this.next_out = 0, this.avail_out = 0, this.total_out = 0, this.msg = "", this.state = null, this.data_type = 2, this.adler = 0;
}
var Ma = xo;
const $a = Object.prototype.toString, {
  Z_NO_FLUSH: To,
  Z_SYNC_FLUSH: Eo,
  Z_FULL_FLUSH: Io,
  Z_FINISH: vo,
  Z_OK: Tr,
  Z_STREAM_END: No,
  Z_DEFAULT_COMPRESSION: Ro,
  Z_DEFAULT_STRATEGY: Co,
  Z_DEFLATED: Po
} = sr;
function lr(t) {
  this.options = Rr.assign({
    level: Ro,
    method: Po,
    chunkSize: 16384,
    windowBits: 15,
    memLevel: 8,
    strategy: Co
  }, t || {});
  let r = this.options;
  r.raw && r.windowBits > 0 ? r.windowBits = -r.windowBits : r.gzip && r.windowBits > 0 && r.windowBits < 16 && (r.windowBits += 16), this.err = 0, this.msg = "", this.ended = !1, this.chunks = [], this.strm = new Ma(), this.strm.avail_out = 0;
  let e = Kt.deflateInit2(
    this.strm,
    r.level,
    r.method,
    r.windowBits,
    r.memLevel,
    r.strategy
  );
  if (e !== Tr)
    throw new Error(At[e]);
  if (r.header && Kt.deflateSetHeader(this.strm, r.header), r.dictionary) {
    let n;
    if (typeof r.dictionary == "string" ? n = nr.string2buf(r.dictionary) : $a.call(r.dictionary) === "[object ArrayBuffer]" ? n = new Uint8Array(r.dictionary) : n = r.dictionary, e = Kt.deflateSetDictionary(this.strm, n), e !== Tr)
      throw new Error(At[e]);
    this._dict_set = !0;
  }
}
lr.prototype.push = function(t, r) {
  const e = this.strm, n = this.options.chunkSize;
  let i, a;
  if (this.ended)
    return !1;
  for (r === ~~r ? a = r : a = r === !0 ? vo : To, typeof t == "string" ? e.input = nr.string2buf(t) : $a.call(t) === "[object ArrayBuffer]" ? e.input = new Uint8Array(t) : e.input = t, e.next_in = 0, e.avail_in = e.input.length; ; ) {
    if (e.avail_out === 0 && (e.output = new Uint8Array(n), e.next_out = 0, e.avail_out = n), (a === Eo || a === Io) && e.avail_out <= 6) {
      this.onData(e.output.subarray(0, e.next_out)), e.avail_out = 0;
      continue;
    }
    if (i = Kt.deflate(e, a), i === No)
      return e.next_out > 0 && this.onData(e.output.subarray(0, e.next_out)), i = Kt.deflateEnd(this.strm), this.onEnd(i), this.ended = !0, i === Tr;
    if (e.avail_out === 0) {
      this.onData(e.output);
      continue;
    }
    if (a > 0 && e.next_out > 0) {
      this.onData(e.output.subarray(0, e.next_out)), e.avail_out = 0;
      continue;
    }
    if (e.avail_in === 0) break;
  }
  return !0;
};
lr.prototype.onData = function(t) {
  this.chunks.push(t);
};
lr.prototype.onEnd = function(t) {
  t === Tr && (this.result = Rr.flattenChunks(this.chunks)), this.chunks = [], this.err = t, this.msg = this.strm.msg;
};
function mn(t, r) {
  const e = new lr(r);
  if (e.push(t, !0), e.err)
    throw e.msg || At[e.err];
  return e.result;
}
function Fo(t, r) {
  return r = r || {}, r.raw = !0, mn(t, r);
}
function ko(t, r) {
  return r = r || {}, r.gzip = !0, mn(t, r);
}
var Lo = lr, Bo = mn, Oo = Fo, Ho = ko, Do = {
  Deflate: Lo,
  deflate: Bo,
  deflateRaw: Oo,
  gzip: Ho
};
const _r = 16209, Mo = 16191;
var $o = function(r, e) {
  let n, i, a, s, o, h, l, f, d, u, c, g, p, y, m, A, E, x, S, R, v, O, C, B;
  const Z = r.state;
  n = r.next_in, C = r.input, i = n + (r.avail_in - 5), a = r.next_out, B = r.output, s = a - (e - r.avail_out), o = a + (r.avail_out - 257), h = Z.dmax, l = Z.wsize, f = Z.whave, d = Z.wnext, u = Z.window, c = Z.hold, g = Z.bits, p = Z.lencode, y = Z.distcode, m = (1 << Z.lenbits) - 1, A = (1 << Z.distbits) - 1;
  e:
    do {
      g < 15 && (c += C[n++] << g, g += 8, c += C[n++] << g, g += 8), E = p[c & m];
      t:
        for (; ; ) {
          if (x = E >>> 24, c >>>= x, g -= x, x = E >>> 16 & 255, x === 0)
            B[a++] = E & 65535;
          else if (x & 16) {
            S = E & 65535, x &= 15, x && (g < x && (c += C[n++] << g, g += 8), S += c & (1 << x) - 1, c >>>= x, g -= x), g < 15 && (c += C[n++] << g, g += 8, c += C[n++] << g, g += 8), E = y[c & A];
            r:
              for (; ; ) {
                if (x = E >>> 24, c >>>= x, g -= x, x = E >>> 16 & 255, x & 16) {
                  if (R = E & 65535, x &= 15, g < x && (c += C[n++] << g, g += 8, g < x && (c += C[n++] << g, g += 8)), R += c & (1 << x) - 1, R > h) {
                    r.msg = "invalid distance too far back", Z.mode = _r;
                    break e;
                  }
                  if (c >>>= x, g -= x, x = a - s, R > x) {
                    if (x = R - x, x > f && Z.sane) {
                      r.msg = "invalid distance too far back", Z.mode = _r;
                      break e;
                    }
                    if (v = 0, O = u, d === 0) {
                      if (v += l - x, x < S) {
                        S -= x;
                        do
                          B[a++] = u[v++];
                        while (--x);
                        v = a - R, O = B;
                      }
                    } else if (d < x) {
                      if (v += l + d - x, x -= d, x < S) {
                        S -= x;
                        do
                          B[a++] = u[v++];
                        while (--x);
                        if (v = 0, d < S) {
                          x = d, S -= x;
                          do
                            B[a++] = u[v++];
                          while (--x);
                          v = a - R, O = B;
                        }
                      }
                    } else if (v += d - x, x < S) {
                      S -= x;
                      do
                        B[a++] = u[v++];
                      while (--x);
                      v = a - R, O = B;
                    }
                    for (; S > 2; )
                      B[a++] = O[v++], B[a++] = O[v++], B[a++] = O[v++], S -= 3;
                    S && (B[a++] = O[v++], S > 1 && (B[a++] = O[v++]));
                  } else {
                    v = a - R;
                    do
                      B[a++] = B[v++], B[a++] = B[v++], B[a++] = B[v++], S -= 3;
                    while (S > 2);
                    S && (B[a++] = B[v++], S > 1 && (B[a++] = B[v++]));
                  }
                } else if (x & 64) {
                  r.msg = "invalid distance code", Z.mode = _r;
                  break e;
                } else {
                  E = y[(E & 65535) + (c & (1 << x) - 1)];
                  continue r;
                }
                break;
              }
          } else if (x & 64)
            if (x & 32) {
              Z.mode = Mo;
              break e;
            } else {
              r.msg = "invalid literal/length code", Z.mode = _r;
              break e;
            }
          else {
            E = p[(E & 65535) + (c & (1 << x) - 1)];
            continue t;
          }
          break;
        }
    } while (n < i && a < o);
  S = g >> 3, n -= S, g -= S << 3, c &= (1 << g) - 1, r.next_in = n, r.next_out = a, r.avail_in = n < i ? 5 + (i - n) : 5 - (n - i), r.avail_out = a < o ? 257 + (o - a) : 257 - (a - o), Z.hold = c, Z.bits = g;
};
const Nt = 15, $n = 852, Un = 592, zn = 0, Gr = 1, Wn = 2, Uo = new Uint16Array([
  /* Length codes 257..285 base */
  3,
  4,
  5,
  6,
  7,
  8,
  9,
  10,
  11,
  13,
  15,
  17,
  19,
  23,
  27,
  31,
  35,
  43,
  51,
  59,
  67,
  83,
  99,
  115,
  131,
  163,
  195,
  227,
  258,
  0,
  0
]), zo = new Uint8Array([
  /* Length codes 257..285 extra */
  16,
  16,
  16,
  16,
  16,
  16,
  16,
  16,
  17,
  17,
  17,
  17,
  18,
  18,
  18,
  18,
  19,
  19,
  19,
  19,
  20,
  20,
  20,
  20,
  21,
  21,
  21,
  21,
  16,
  72,
  78
]), Wo = new Uint16Array([
  /* Distance codes 0..29 base */
  1,
  2,
  3,
  4,
  5,
  7,
  9,
  13,
  17,
  25,
  33,
  49,
  65,
  97,
  129,
  193,
  257,
  385,
  513,
  769,
  1025,
  1537,
  2049,
  3073,
  4097,
  6145,
  8193,
  12289,
  16385,
  24577,
  0,
  0
]), jo = new Uint8Array([
  /* Distance codes 0..29 extra */
  16,
  16,
  16,
  16,
  17,
  17,
  18,
  18,
  19,
  19,
  20,
  20,
  21,
  21,
  22,
  22,
  23,
  23,
  24,
  24,
  25,
  25,
  26,
  26,
  27,
  27,
  28,
  28,
  29,
  29,
  64,
  64
]), Go = (t, r, e, n, i, a, s, o) => {
  const h = o.bits;
  let l = 0, f = 0, d = 0, u = 0, c = 0, g = 0, p = 0, y = 0, m = 0, A = 0, E, x, S, R, v, O = null, C;
  const B = new Uint16Array(Nt + 1), Z = new Uint16Array(Nt + 1);
  let T = null, j, _, K;
  for (l = 0; l <= Nt; l++)
    B[l] = 0;
  for (f = 0; f < n; f++)
    B[r[e + f]]++;
  for (c = h, u = Nt; u >= 1 && B[u] === 0; u--)
    ;
  if (c > u && (c = u), u === 0)
    return i[a++] = 1 << 24 | 64 << 16 | 0, i[a++] = 1 << 24 | 64 << 16 | 0, o.bits = 1, 0;
  for (d = 1; d < u && B[d] === 0; d++)
    ;
  for (c < d && (c = d), y = 1, l = 1; l <= Nt; l++)
    if (y <<= 1, y -= B[l], y < 0)
      return -1;
  if (y > 0 && (t === zn || u !== 1))
    return -1;
  for (Z[1] = 0, l = 1; l < Nt; l++)
    Z[l + 1] = Z[l] + B[l];
  for (f = 0; f < n; f++)
    r[e + f] !== 0 && (s[Z[r[e + f]]++] = f);
  if (t === zn ? (O = T = s, C = 20) : t === Gr ? (O = Uo, T = zo, C = 257) : (O = Wo, T = jo, C = 0), A = 0, f = 0, l = d, v = a, g = c, p = 0, S = -1, m = 1 << c, R = m - 1, t === Gr && m > $n || t === Wn && m > Un)
    return 1;
  for (; ; ) {
    j = l - p, s[f] + 1 < C ? (_ = 0, K = s[f]) : s[f] >= C ? (_ = T[s[f] - C], K = O[s[f] - C]) : (_ = 96, K = 0), E = 1 << l - p, x = 1 << g, d = x;
    do
      x -= E, i[v + (A >> p) + x] = j << 24 | _ << 16 | K | 0;
    while (x !== 0);
    for (E = 1 << l - 1; A & E; )
      E >>= 1;
    if (E !== 0 ? (A &= E - 1, A += E) : A = 0, f++, --B[l] === 0) {
      if (l === u)
        break;
      l = r[e + s[f]];
    }
    if (l > c && (A & R) !== S) {
      for (p === 0 && (p = c), v += d, g = l - p, y = 1 << g; g + p < u && (y -= B[g + p], !(y <= 0)); )
        g++, y <<= 1;
      if (m += 1 << g, t === Gr && m > $n || t === Wn && m > Un)
        return 1;
      S = A & R, i[S] = c << 24 | g << 16 | v - a | 0;
    }
  }
  return A !== 0 && (i[v + A] = l - p << 24 | 64 << 16 | 0), o.bits = c, 0;
};
var Jt = Go;
const Xo = 0, Ua = 1, za = 2, {
  Z_FINISH: jn,
  Z_BLOCK: Zo,
  Z_TREES: yr,
  Z_OK: bt,
  Z_STREAM_END: Ko,
  Z_NEED_DICT: Jo,
  Z_STREAM_ERROR: Ke,
  Z_DATA_ERROR: Wa,
  Z_MEM_ERROR: ja,
  Z_BUF_ERROR: Vo,
  Z_DEFLATED: Gn
} = sr, Cr = 16180, Xn = 16181, Zn = 16182, Kn = 16183, Jn = 16184, Vn = 16185, Yn = 16186, qn = 16187, Qn = 16188, ea = 16189, Er = 16190, it = 16191, Xr = 16192, ta = 16193, Zr = 16194, ra = 16195, na = 16196, aa = 16197, ia = 16198, Ar = 16199, Sr = 16200, sa = 16201, oa = 16202, la = 16203, ha = 16204, ca = 16205, Kr = 16206, fa = 16207, da = 16208, Pe = 16209, Ga = 16210, Xa = 16211, Yo = 852, qo = 592, Qo = 15, el = Qo, pa = (t) => (t >>> 24 & 255) + (t >>> 8 & 65280) + ((t & 65280) << 8) + ((t & 255) << 24);
function tl() {
  this.strm = null, this.mode = 0, this.last = !1, this.wrap = 0, this.havedict = !1, this.flags = 0, this.dmax = 0, this.check = 0, this.total = 0, this.head = null, this.wbits = 0, this.wsize = 0, this.whave = 0, this.wnext = 0, this.window = null, this.hold = 0, this.bits = 0, this.length = 0, this.offset = 0, this.extra = 0, this.lencode = null, this.distcode = null, this.lenbits = 0, this.distbits = 0, this.ncode = 0, this.nlen = 0, this.ndist = 0, this.have = 0, this.next = null, this.lens = new Uint16Array(320), this.work = new Uint16Array(288), this.lendyn = null, this.distdyn = null, this.sane = 0, this.back = 0, this.was = 0;
}
const xt = (t) => {
  if (!t)
    return 1;
  const r = t.state;
  return !r || r.strm !== t || r.mode < Cr || r.mode > Xa ? 1 : 0;
}, Za = (t) => {
  if (xt(t))
    return Ke;
  const r = t.state;
  return t.total_in = t.total_out = r.total = 0, t.msg = "", r.wrap && (t.adler = r.wrap & 1), r.mode = Cr, r.last = 0, r.havedict = 0, r.flags = -1, r.dmax = 32768, r.head = null, r.hold = 0, r.bits = 0, r.lencode = r.lendyn = new Int32Array(Yo), r.distcode = r.distdyn = new Int32Array(qo), r.sane = 1, r.back = -1, bt;
}, Ka = (t) => {
  if (xt(t))
    return Ke;
  const r = t.state;
  return r.wsize = 0, r.whave = 0, r.wnext = 0, Za(t);
}, Ja = (t, r) => {
  let e;
  if (xt(t))
    return Ke;
  const n = t.state;
  return r < 0 ? (e = 0, r = -r) : (e = (r >> 4) + 5, r < 48 && (r &= 15)), r && (r < 8 || r > 15) ? Ke : (n.window !== null && n.wbits !== r && (n.window = null), n.wrap = e, n.wbits = r, Ka(t));
}, Va = (t, r) => {
  if (!t)
    return Ke;
  const e = new tl();
  t.state = e, e.strm = t, e.window = null, e.mode = Cr;
  const n = Ja(t, r);
  return n !== bt && (t.state = null), n;
}, rl = (t) => Va(t, el);
let ua = !0, Jr, Vr;
const nl = (t) => {
  if (ua) {
    Jr = new Int32Array(512), Vr = new Int32Array(32);
    let r = 0;
    for (; r < 144; )
      t.lens[r++] = 8;
    for (; r < 256; )
      t.lens[r++] = 9;
    for (; r < 280; )
      t.lens[r++] = 7;
    for (; r < 288; )
      t.lens[r++] = 8;
    for (Jt(Ua, t.lens, 0, 288, Jr, 0, t.work, { bits: 9 }), r = 0; r < 32; )
      t.lens[r++] = 5;
    Jt(za, t.lens, 0, 32, Vr, 0, t.work, { bits: 5 }), ua = !1;
  }
  t.lencode = Jr, t.lenbits = 9, t.distcode = Vr, t.distbits = 5;
}, Ya = (t, r, e, n) => {
  let i;
  const a = t.state;
  return a.window === null && (a.wsize = 1 << a.wbits, a.wnext = 0, a.whave = 0, a.window = new Uint8Array(a.wsize)), n >= a.wsize ? (a.window.set(r.subarray(e - a.wsize, e), 0), a.wnext = 0, a.whave = a.wsize) : (i = a.wsize - a.wnext, i > n && (i = n), a.window.set(r.subarray(e - n, e - n + i), a.wnext), n -= i, n ? (a.window.set(r.subarray(e - n, e), 0), a.wnext = n, a.whave = a.wsize) : (a.wnext += i, a.wnext === a.wsize && (a.wnext = 0), a.whave < a.wsize && (a.whave += i))), 0;
}, al = (t, r) => {
  let e, n, i, a, s, o, h, l, f, d, u, c, g, p, y = 0, m, A, E, x, S, R, v, O;
  const C = new Uint8Array(4);
  let B, Z;
  const T = (
    /* permutation of code lengths */
    new Uint8Array([16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15])
  );
  if (xt(t) || !t.output || !t.input && t.avail_in !== 0)
    return Ke;
  e = t.state, e.mode === it && (e.mode = Xr), s = t.next_out, i = t.output, h = t.avail_out, a = t.next_in, n = t.input, o = t.avail_in, l = e.hold, f = e.bits, d = o, u = h, O = bt;
  e:
    for (; ; )
      switch (e.mode) {
        case Cr:
          if (e.wrap === 0) {
            e.mode = Xr;
            break;
          }
          for (; f < 16; ) {
            if (o === 0)
              break e;
            o--, l += n[a++] << f, f += 8;
          }
          if (e.wrap & 2 && l === 35615) {
            e.wbits === 0 && (e.wbits = 15), e.check = 0, C[0] = l & 255, C[1] = l >>> 8 & 255, e.check = Le(e.check, C, 2, 0), l = 0, f = 0, e.mode = Xn;
            break;
          }
          if (e.head && (e.head.done = !1), !(e.wrap & 1) || /* check if zlib header allowed */
          (((l & 255) << 8) + (l >> 8)) % 31) {
            t.msg = "incorrect header check", e.mode = Pe;
            break;
          }
          if ((l & 15) !== Gn) {
            t.msg = "unknown compression method", e.mode = Pe;
            break;
          }
          if (l >>>= 4, f -= 4, v = (l & 15) + 8, e.wbits === 0 && (e.wbits = v), v > 15 || v > e.wbits) {
            t.msg = "invalid window size", e.mode = Pe;
            break;
          }
          e.dmax = 1 << e.wbits, e.flags = 0, t.adler = e.check = 1, e.mode = l & 512 ? ea : it, l = 0, f = 0;
          break;
        case Xn:
          for (; f < 16; ) {
            if (o === 0)
              break e;
            o--, l += n[a++] << f, f += 8;
          }
          if (e.flags = l, (e.flags & 255) !== Gn) {
            t.msg = "unknown compression method", e.mode = Pe;
            break;
          }
          if (e.flags & 57344) {
            t.msg = "unknown header flags set", e.mode = Pe;
            break;
          }
          e.head && (e.head.text = l >> 8 & 1), e.flags & 512 && e.wrap & 4 && (C[0] = l & 255, C[1] = l >>> 8 & 255, e.check = Le(e.check, C, 2, 0)), l = 0, f = 0, e.mode = Zn;
        case Zn:
          for (; f < 32; ) {
            if (o === 0)
              break e;
            o--, l += n[a++] << f, f += 8;
          }
          e.head && (e.head.time = l), e.flags & 512 && e.wrap & 4 && (C[0] = l & 255, C[1] = l >>> 8 & 255, C[2] = l >>> 16 & 255, C[3] = l >>> 24 & 255, e.check = Le(e.check, C, 4, 0)), l = 0, f = 0, e.mode = Kn;
        case Kn:
          for (; f < 16; ) {
            if (o === 0)
              break e;
            o--, l += n[a++] << f, f += 8;
          }
          e.head && (e.head.xflags = l & 255, e.head.os = l >> 8), e.flags & 512 && e.wrap & 4 && (C[0] = l & 255, C[1] = l >>> 8 & 255, e.check = Le(e.check, C, 2, 0)), l = 0, f = 0, e.mode = Jn;
        case Jn:
          if (e.flags & 1024) {
            for (; f < 16; ) {
              if (o === 0)
                break e;
              o--, l += n[a++] << f, f += 8;
            }
            e.length = l, e.head && (e.head.extra_len = l), e.flags & 512 && e.wrap & 4 && (C[0] = l & 255, C[1] = l >>> 8 & 255, e.check = Le(e.check, C, 2, 0)), l = 0, f = 0;
          } else e.head && (e.head.extra = null);
          e.mode = Vn;
        case Vn:
          if (e.flags & 1024 && (c = e.length, c > o && (c = o), c && (e.head && (v = e.head.extra_len - e.length, e.head.extra || (e.head.extra = new Uint8Array(e.head.extra_len)), e.head.extra.set(
            n.subarray(
              a,
              // extra field is limited to 65536 bytes
              // - no need for additional size check
              a + c
            ),
            /*len + copy > state.head.extra_max - len ? state.head.extra_max : copy,*/
            v
          )), e.flags & 512 && e.wrap & 4 && (e.check = Le(e.check, n, c, a)), o -= c, a += c, e.length -= c), e.length))
            break e;
          e.length = 0, e.mode = Yn;
        case Yn:
          if (e.flags & 2048) {
            if (o === 0)
              break e;
            c = 0;
            do
              v = n[a + c++], e.head && v && e.length < 65536 && (e.head.name += String.fromCharCode(v));
            while (v && c < o);
            if (e.flags & 512 && e.wrap & 4 && (e.check = Le(e.check, n, c, a)), o -= c, a += c, v)
              break e;
          } else e.head && (e.head.name = null);
          e.length = 0, e.mode = qn;
        case qn:
          if (e.flags & 4096) {
            if (o === 0)
              break e;
            c = 0;
            do
              v = n[a + c++], e.head && v && e.length < 65536 && (e.head.comment += String.fromCharCode(v));
            while (v && c < o);
            if (e.flags & 512 && e.wrap & 4 && (e.check = Le(e.check, n, c, a)), o -= c, a += c, v)
              break e;
          } else e.head && (e.head.comment = null);
          e.mode = Qn;
        case Qn:
          if (e.flags & 512) {
            for (; f < 16; ) {
              if (o === 0)
                break e;
              o--, l += n[a++] << f, f += 8;
            }
            if (e.wrap & 4 && l !== (e.check & 65535)) {
              t.msg = "header crc mismatch", e.mode = Pe;
              break;
            }
            l = 0, f = 0;
          }
          e.head && (e.head.hcrc = e.flags >> 9 & 1, e.head.done = !0), t.adler = e.check = 0, e.mode = it;
          break;
        case ea:
          for (; f < 32; ) {
            if (o === 0)
              break e;
            o--, l += n[a++] << f, f += 8;
          }
          t.adler = e.check = pa(l), l = 0, f = 0, e.mode = Er;
        case Er:
          if (e.havedict === 0)
            return t.next_out = s, t.avail_out = h, t.next_in = a, t.avail_in = o, e.hold = l, e.bits = f, Jo;
          t.adler = e.check = 1, e.mode = it;
        case it:
          if (r === Zo || r === yr)
            break e;
        case Xr:
          if (e.last) {
            l >>>= f & 7, f -= f & 7, e.mode = Kr;
            break;
          }
          for (; f < 3; ) {
            if (o === 0)
              break e;
            o--, l += n[a++] << f, f += 8;
          }
          switch (e.last = l & 1, l >>>= 1, f -= 1, l & 3) {
            case 0:
              e.mode = ta;
              break;
            case 1:
              if (nl(e), e.mode = Ar, r === yr) {
                l >>>= 2, f -= 2;
                break e;
              }
              break;
            case 2:
              e.mode = na;
              break;
            case 3:
              t.msg = "invalid block type", e.mode = Pe;
          }
          l >>>= 2, f -= 2;
          break;
        case ta:
          for (l >>>= f & 7, f -= f & 7; f < 32; ) {
            if (o === 0)
              break e;
            o--, l += n[a++] << f, f += 8;
          }
          if ((l & 65535) !== (l >>> 16 ^ 65535)) {
            t.msg = "invalid stored block lengths", e.mode = Pe;
            break;
          }
          if (e.length = l & 65535, l = 0, f = 0, e.mode = Zr, r === yr)
            break e;
        case Zr:
          e.mode = ra;
        case ra:
          if (c = e.length, c) {
            if (c > o && (c = o), c > h && (c = h), c === 0)
              break e;
            i.set(n.subarray(a, a + c), s), o -= c, a += c, h -= c, s += c, e.length -= c;
            break;
          }
          e.mode = it;
          break;
        case na:
          for (; f < 14; ) {
            if (o === 0)
              break e;
            o--, l += n[a++] << f, f += 8;
          }
          if (e.nlen = (l & 31) + 257, l >>>= 5, f -= 5, e.ndist = (l & 31) + 1, l >>>= 5, f -= 5, e.ncode = (l & 15) + 4, l >>>= 4, f -= 4, e.nlen > 286 || e.ndist > 30) {
            t.msg = "too many length or distance symbols", e.mode = Pe;
            break;
          }
          e.have = 0, e.mode = aa;
        case aa:
          for (; e.have < e.ncode; ) {
            for (; f < 3; ) {
              if (o === 0)
                break e;
              o--, l += n[a++] << f, f += 8;
            }
            e.lens[T[e.have++]] = l & 7, l >>>= 3, f -= 3;
          }
          for (; e.have < 19; )
            e.lens[T[e.have++]] = 0;
          if (e.lencode = e.lendyn, e.lenbits = 7, B = { bits: e.lenbits }, O = Jt(Xo, e.lens, 0, 19, e.lencode, 0, e.work, B), e.lenbits = B.bits, O) {
            t.msg = "invalid code lengths set", e.mode = Pe;
            break;
          }
          e.have = 0, e.mode = ia;
        case ia:
          for (; e.have < e.nlen + e.ndist; ) {
            for (; y = e.lencode[l & (1 << e.lenbits) - 1], m = y >>> 24, A = y >>> 16 & 255, E = y & 65535, !(m <= f); ) {
              if (o === 0)
                break e;
              o--, l += n[a++] << f, f += 8;
            }
            if (E < 16)
              l >>>= m, f -= m, e.lens[e.have++] = E;
            else {
              if (E === 16) {
                for (Z = m + 2; f < Z; ) {
                  if (o === 0)
                    break e;
                  o--, l += n[a++] << f, f += 8;
                }
                if (l >>>= m, f -= m, e.have === 0) {
                  t.msg = "invalid bit length repeat", e.mode = Pe;
                  break;
                }
                v = e.lens[e.have - 1], c = 3 + (l & 3), l >>>= 2, f -= 2;
              } else if (E === 17) {
                for (Z = m + 3; f < Z; ) {
                  if (o === 0)
                    break e;
                  o--, l += n[a++] << f, f += 8;
                }
                l >>>= m, f -= m, v = 0, c = 3 + (l & 7), l >>>= 3, f -= 3;
              } else {
                for (Z = m + 7; f < Z; ) {
                  if (o === 0)
                    break e;
                  o--, l += n[a++] << f, f += 8;
                }
                l >>>= m, f -= m, v = 0, c = 11 + (l & 127), l >>>= 7, f -= 7;
              }
              if (e.have + c > e.nlen + e.ndist) {
                t.msg = "invalid bit length repeat", e.mode = Pe;
                break;
              }
              for (; c--; )
                e.lens[e.have++] = v;
            }
          }
          if (e.mode === Pe)
            break;
          if (e.lens[256] === 0) {
            t.msg = "invalid code -- missing end-of-block", e.mode = Pe;
            break;
          }
          if (e.lenbits = 9, B = { bits: e.lenbits }, O = Jt(Ua, e.lens, 0, e.nlen, e.lencode, 0, e.work, B), e.lenbits = B.bits, O) {
            t.msg = "invalid literal/lengths set", e.mode = Pe;
            break;
          }
          if (e.distbits = 6, e.distcode = e.distdyn, B = { bits: e.distbits }, O = Jt(za, e.lens, e.nlen, e.ndist, e.distcode, 0, e.work, B), e.distbits = B.bits, O) {
            t.msg = "invalid distances set", e.mode = Pe;
            break;
          }
          if (e.mode = Ar, r === yr)
            break e;
        case Ar:
          e.mode = Sr;
        case Sr:
          if (o >= 6 && h >= 258) {
            t.next_out = s, t.avail_out = h, t.next_in = a, t.avail_in = o, e.hold = l, e.bits = f, $o(t, u), s = t.next_out, i = t.output, h = t.avail_out, a = t.next_in, n = t.input, o = t.avail_in, l = e.hold, f = e.bits, e.mode === it && (e.back = -1);
            break;
          }
          for (e.back = 0; y = e.lencode[l & (1 << e.lenbits) - 1], m = y >>> 24, A = y >>> 16 & 255, E = y & 65535, !(m <= f); ) {
            if (o === 0)
              break e;
            o--, l += n[a++] << f, f += 8;
          }
          if (A && !(A & 240)) {
            for (x = m, S = A, R = E; y = e.lencode[R + ((l & (1 << x + S) - 1) >> x)], m = y >>> 24, A = y >>> 16 & 255, E = y & 65535, !(x + m <= f); ) {
              if (o === 0)
                break e;
              o--, l += n[a++] << f, f += 8;
            }
            l >>>= x, f -= x, e.back += x;
          }
          if (l >>>= m, f -= m, e.back += m, e.length = E, A === 0) {
            e.mode = ca;
            break;
          }
          if (A & 32) {
            e.back = -1, e.mode = it;
            break;
          }
          if (A & 64) {
            t.msg = "invalid literal/length code", e.mode = Pe;
            break;
          }
          e.extra = A & 15, e.mode = sa;
        case sa:
          if (e.extra) {
            for (Z = e.extra; f < Z; ) {
              if (o === 0)
                break e;
              o--, l += n[a++] << f, f += 8;
            }
            e.length += l & (1 << e.extra) - 1, l >>>= e.extra, f -= e.extra, e.back += e.extra;
          }
          e.was = e.length, e.mode = oa;
        case oa:
          for (; y = e.distcode[l & (1 << e.distbits) - 1], m = y >>> 24, A = y >>> 16 & 255, E = y & 65535, !(m <= f); ) {
            if (o === 0)
              break e;
            o--, l += n[a++] << f, f += 8;
          }
          if (!(A & 240)) {
            for (x = m, S = A, R = E; y = e.distcode[R + ((l & (1 << x + S) - 1) >> x)], m = y >>> 24, A = y >>> 16 & 255, E = y & 65535, !(x + m <= f); ) {
              if (o === 0)
                break e;
              o--, l += n[a++] << f, f += 8;
            }
            l >>>= x, f -= x, e.back += x;
          }
          if (l >>>= m, f -= m, e.back += m, A & 64) {
            t.msg = "invalid distance code", e.mode = Pe;
            break;
          }
          e.offset = E, e.extra = A & 15, e.mode = la;
        case la:
          if (e.extra) {
            for (Z = e.extra; f < Z; ) {
              if (o === 0)
                break e;
              o--, l += n[a++] << f, f += 8;
            }
            e.offset += l & (1 << e.extra) - 1, l >>>= e.extra, f -= e.extra, e.back += e.extra;
          }
          if (e.offset > e.dmax) {
            t.msg = "invalid distance too far back", e.mode = Pe;
            break;
          }
          e.mode = ha;
        case ha:
          if (h === 0)
            break e;
          if (c = u - h, e.offset > c) {
            if (c = e.offset - c, c > e.whave && e.sane) {
              t.msg = "invalid distance too far back", e.mode = Pe;
              break;
            }
            c > e.wnext ? (c -= e.wnext, g = e.wsize - c) : g = e.wnext - c, c > e.length && (c = e.length), p = e.window;
          } else
            p = i, g = s - e.offset, c = e.length;
          c > h && (c = h), h -= c, e.length -= c;
          do
            i[s++] = p[g++];
          while (--c);
          e.length === 0 && (e.mode = Sr);
          break;
        case ca:
          if (h === 0)
            break e;
          i[s++] = e.length, h--, e.mode = Sr;
          break;
        case Kr:
          if (e.wrap) {
            for (; f < 32; ) {
              if (o === 0)
                break e;
              o--, l |= n[a++] << f, f += 8;
            }
            if (u -= h, t.total_out += u, e.total += u, e.wrap & 4 && u && (t.adler = e.check = /*UPDATE_CHECK(state.check, put - _out, _out);*/
            e.flags ? Le(e.check, i, u, s - u) : tr(e.check, i, u, s - u)), u = h, e.wrap & 4 && (e.flags ? l : pa(l)) !== e.check) {
              t.msg = "incorrect data check", e.mode = Pe;
              break;
            }
            l = 0, f = 0;
          }
          e.mode = fa;
        case fa:
          if (e.wrap && e.flags) {
            for (; f < 32; ) {
              if (o === 0)
                break e;
              o--, l += n[a++] << f, f += 8;
            }
            if (e.wrap & 4 && l !== (e.total & 4294967295)) {
              t.msg = "incorrect length check", e.mode = Pe;
              break;
            }
            l = 0, f = 0;
          }
          e.mode = da;
        case da:
          O = Ko;
          break e;
        case Pe:
          O = Wa;
          break e;
        case Ga:
          return ja;
        case Xa:
        default:
          return Ke;
      }
  return t.next_out = s, t.avail_out = h, t.next_in = a, t.avail_in = o, e.hold = l, e.bits = f, (e.wsize || u !== t.avail_out && e.mode < Pe && (e.mode < Kr || r !== jn)) && Ya(t, t.output, t.next_out, u - t.avail_out), d -= t.avail_in, u -= t.avail_out, t.total_in += d, t.total_out += u, e.total += u, e.wrap & 4 && u && (t.adler = e.check = /*UPDATE_CHECK(state.check, strm.next_out - _out, _out);*/
  e.flags ? Le(e.check, i, u, t.next_out - u) : tr(e.check, i, u, t.next_out - u)), t.data_type = e.bits + (e.last ? 64 : 0) + (e.mode === it ? 128 : 0) + (e.mode === Ar || e.mode === Zr ? 256 : 0), (d === 0 && u === 0 || r === jn) && O === bt && (O = Vo), O;
}, il = (t) => {
  if (xt(t))
    return Ke;
  let r = t.state;
  return r.window && (r.window = null), t.state = null, bt;
}, sl = (t, r) => {
  if (xt(t))
    return Ke;
  const e = t.state;
  return e.wrap & 2 ? (e.head = r, r.done = !1, bt) : Ke;
}, ol = (t, r) => {
  const e = r.length;
  let n, i, a;
  return xt(t) || (n = t.state, n.wrap !== 0 && n.mode !== Er) ? Ke : n.mode === Er && (i = 1, i = tr(i, r, e, 0), i !== n.check) ? Wa : (a = Ya(t, r, e, e), a ? (n.mode = Ga, ja) : (n.havedict = 1, bt));
};
var ll = Ka, hl = Ja, cl = Za, fl = rl, dl = Va, pl = al, ul = il, gl = sl, ml = ol, wl = "pako inflate (from Nodeca project)", ot = {
  inflateReset: ll,
  inflateReset2: hl,
  inflateResetKeep: cl,
  inflateInit: fl,
  inflateInit2: dl,
  inflate: pl,
  inflateEnd: ul,
  inflateGetHeader: gl,
  inflateSetDictionary: ml,
  inflateInfo: wl
};
function _l() {
  this.text = 0, this.time = 0, this.xflags = 0, this.os = 0, this.extra = null, this.extra_len = 0, this.name = "", this.comment = "", this.hcrc = 0, this.done = !1;
}
var yl = _l;
const qa = Object.prototype.toString, {
  Z_NO_FLUSH: Al,
  Z_FINISH: Sl,
  Z_OK: ar,
  Z_STREAM_END: Yr,
  Z_NEED_DICT: qr,
  Z_STREAM_ERROR: bl,
  Z_DATA_ERROR: ga,
  Z_MEM_ERROR: xl
} = sr;
function hr(t) {
  this.options = Rr.assign({
    chunkSize: 1024 * 64,
    windowBits: 15,
    to: ""
  }, t || {});
  const r = this.options;
  r.raw && r.windowBits >= 0 && r.windowBits < 16 && (r.windowBits = -r.windowBits, r.windowBits === 0 && (r.windowBits = -15)), r.windowBits >= 0 && r.windowBits < 16 && !(t && t.windowBits) && (r.windowBits += 32), r.windowBits > 15 && r.windowBits < 48 && (r.windowBits & 15 || (r.windowBits |= 15)), this.err = 0, this.msg = "", this.ended = !1, this.chunks = [], this.strm = new Ma(), this.strm.avail_out = 0;
  let e = ot.inflateInit2(
    this.strm,
    r.windowBits
  );
  if (e !== ar)
    throw new Error(At[e]);
  if (this.header = new yl(), ot.inflateGetHeader(this.strm, this.header), r.dictionary && (typeof r.dictionary == "string" ? r.dictionary = nr.string2buf(r.dictionary) : qa.call(r.dictionary) === "[object ArrayBuffer]" && (r.dictionary = new Uint8Array(r.dictionary)), r.raw && (e = ot.inflateSetDictionary(this.strm, r.dictionary), e !== ar)))
    throw new Error(At[e]);
}
hr.prototype.push = function(t, r) {
  const e = this.strm, n = this.options.chunkSize, i = this.options.dictionary;
  let a, s, o;
  if (this.ended) return !1;
  for (r === ~~r ? s = r : s = r === !0 ? Sl : Al, qa.call(t) === "[object ArrayBuffer]" ? e.input = new Uint8Array(t) : e.input = t, e.next_in = 0, e.avail_in = e.input.length; ; ) {
    for (e.avail_out === 0 && (e.output = new Uint8Array(n), e.next_out = 0, e.avail_out = n), a = ot.inflate(e, s), a === qr && i && (a = ot.inflateSetDictionary(e, i), a === ar ? a = ot.inflate(e, s) : a === ga && (a = qr)); e.avail_in > 0 && a === Yr && e.state.wrap > 0 && t[e.next_in] !== 0; )
      ot.inflateReset(e), a = ot.inflate(e, s);
    switch (a) {
      case bl:
      case ga:
      case qr:
      case xl:
        return this.onEnd(a), this.ended = !0, !1;
    }
    if (o = e.avail_out, e.next_out && (e.avail_out === 0 || a === Yr))
      if (this.options.to === "string") {
        let h = nr.utf8border(e.output, e.next_out), l = e.next_out - h, f = nr.buf2string(e.output, h);
        e.next_out = l, e.avail_out = n - l, l && e.output.set(e.output.subarray(h, h + l), 0), this.onData(f);
      } else
        this.onData(e.output.length === e.next_out ? e.output : e.output.subarray(0, e.next_out));
    if (!(a === ar && o === 0)) {
      if (a === Yr)
        return a = ot.inflateEnd(this.strm), this.onEnd(a), this.ended = !0, !0;
      if (e.avail_in === 0) break;
    }
  }
  return !0;
};
hr.prototype.onData = function(t) {
  this.chunks.push(t);
};
hr.prototype.onEnd = function(t) {
  t === ar && (this.options.to === "string" ? this.result = this.chunks.join("") : this.result = Rr.flattenChunks(this.chunks)), this.chunks = [], this.err = t, this.msg = this.strm.msg;
};
function wn(t, r) {
  const e = new hr(r);
  if (e.push(t), e.err) throw e.msg || At[e.err];
  return e.result;
}
function Tl(t, r) {
  return r = r || {}, r.raw = !0, wn(t, r);
}
var El = hr, Il = wn, vl = Tl, Nl = wn, Rl = {
  Inflate: El,
  inflate: Il,
  inflateRaw: vl,
  ungzip: Nl
};
const { Deflate: Cl, deflate: Pl, deflateRaw: Fl, gzip: kl } = Do, { Inflate: Ll, inflate: Bl, inflateRaw: Ol, ungzip: Hl } = Rl;
var Dl = Cl, Ml = Pl, $l = Fl, Ul = kl, zl = Ll, Wl = Bl, jl = Ol, Gl = Hl, Xl = sr, Vt = {
  Deflate: Dl,
  deflate: Ml,
  deflateRaw: $l,
  gzip: Ul,
  Inflate: zl,
  inflate: Wl,
  inflateRaw: jl,
  ungzip: Gl,
  constants: Xl
};
function zt(t) {
  let r = "", e = 0;
  const n = Math.floor(t.length / 2), i = new DataView(t.buffer, t.byteOffset, t.byteLength);
  for (; e < n; ) {
    const a = i.getUint16(e * 2, !0);
    a === 10 || a === 13 ? (r += `
`, e += 1) : a >= 0 && a <= 31 ? [0, 10, 13, 24, 25, 26, 27, 28, 29, 30, 31].includes(a) ? e += 1 : [1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23].includes(a) ? e += 8 : e += 1 : (r += String.fromCharCode(a), e += 1);
  }
  return r.trim();
}
function Wt(t, r, e) {
  let n = "";
  const i = new DataView(t.buffer, t.byteOffset, t.byteLength), a = Math.floor(t.length / 2);
  let s = 0, o = { isBold: !1, isItalic: !1, isUnderline: !1, isStrike: !1 }, h = 0;
  const l = (d) => {
    let u = "";
    return d.baseSize && d.baseSize !== 10 && (u += `<span style="font-size: ${d.baseSize}pt;">`), d.isBold && (u += "<b>"), d.isItalic && (u += "<i>"), d.isUnderline && (u += "<u>"), d.isStrike && (u += "<del>"), u;
  }, f = (d) => {
    let u = "";
    return d.isStrike && (u += "</del>"), d.isUnderline && (u += "</u>"), d.isItalic && (u += "</i>"), d.isBold && (u += "</b>"), d.baseSize && d.baseSize !== 10 && (u += "</span>"), u;
  };
  for (; h < a; ) {
    if (s < r.length && r[s].pos === h) {
      const g = r[s].shapeId, p = g < e.length ? e[g] : { isBold: !1 };
      JSON.stringify(o) !== JSON.stringify(p) && (n += f(o), n += l(p), o = p), s++;
    }
    const d = i.getUint16(h * 2, !0);
    let u = "", c = 1;
    d === 10 || d === 13 ? u = `
` : d >= 0 && d <= 31 ? [0, 10, 13, 24, 25, 26, 27, 28, 29, 30, 31].includes(d) ? c = 1 : [1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23].includes(d) ? c = 8 : c = 1 : u = String.fromCharCode(d), u && (n += u), h += c;
  }
  return n += f(o), n = n.replace(/<del><\/del>/g, ""), n = n.replace(/<u><\/u>/g, ""), n = n.replace(/<i><\/i>/g, ""), n = n.replace(/<b><\/b>/g, ""), n = n.replace(/<span[^>]*><\/span>/g, ""), n.trim();
}
let Zl = class {
  constructor(r) {
    ce(this, "ole");
    ce(this, "charShapes", []);
    ce(this, "paraShapes", []);
    ce(this, "borderFills", [
      {
        left: { type: 0, thickness: 0, color: "#000000" },
        right: { type: 0, thickness: 0, color: "#000000" },
        top: { type: 0, thickness: 0, color: "#000000" },
        bottom: { type: 0, thickness: 0, color: "#000000" }
      }
    ]);
    // ID 0 is null
    ce(this, "compressed", !1);
    this.ole = lt.read(r, { type: "buffer" });
  }
  parse() {
    return this.parseFileHeader(), this.parseDocInfo(), this.parseBodyText();
  }
  parseFileHeader() {
    const r = this.ole.FullPaths.find(
      (s) => s.toLowerCase().endsWith("fileheader")
    );
    if (!r)
      throw new Error("Not a valid HWP 5.0 file (FileHeader not found)");
    const e = lt.find(this.ole, r);
    if (!e || !e.content || e.content.length < 40)
      throw new Error("Invalid FileHeader");
    const n = new Uint8Array(e.content), a = new DataView(
      n.buffer,
      n.byteOffset,
      n.byteLength
    ).getUint32(36, !0);
    this.compressed = (a & 1) !== 0;
  }
  parseDocInfo() {
    const r = this.ole.FullPaths.find(
      (s) => s.toLowerCase().endsWith("docinfo")
    );
    if (!r) return;
    const e = lt.find(this.ole, r);
    if (!e || !e.content) return;
    let n = new Uint8Array(e.content);
    if (this.compressed)
      try {
        n = Vt.inflate(n, { raw: !0 });
      } catch {
      }
    const i = new DataView(
      n.buffer,
      n.byteOffset,
      n.byteLength
    );
    let a = 0;
    for (; a < n.length && !(a + 4 > n.length); ) {
      const s = i.getUint32(a, !0);
      a += 4;
      const o = s & 1023;
      let h = s >>> 20 & 4095;
      if (h === 4095) {
        if (a + 4 > n.length) break;
        h = i.getUint32(a, !0), a += 4;
      }
      if (o === 18) {
        if (a + 60 <= n.length) {
          const l = i.getUint32(a + 36, !0), f = (l & 1) !== 0, d = (l & 2) !== 0, u = (l & 4) !== 0, c = (l & 8) !== 0, g = i.getUint32(a + 56, !0), p = Math.round(g / 100);
          this.charShapes.push({
            isBold: d,
            isItalic: f,
            isUnderline: u,
            isStrike: c,
            baseSize: p
          });
        } else if (a + 40 <= n.length) {
          const f = (i.getUint32(a + 36, !0) & 2) !== 0;
          this.charShapes.push({ isBold: f });
        }
      } else if (o === 22) {
        if (a + 4 <= n.length) {
          const f = i.getUint32(a, !0) >>> 2 & 7;
          this.paraShapes.push({ alignment: f });
        }
      } else if (o === 19) {
        let l;
        const f = (p) => {
          if (a + p + 8 > n.length)
            return { type: 0, thickness: 0, color: "#000000" };
          const y = i.getUint16(a + p, !0), m = i.getUint16(a + p + 2, !0), A = i.getUint32(a + p + 4, !0);
          return {
            type: y,
            thickness: m,
            color: `#${(A & 255).toString(16).padStart(2, "0")}${(A >> 8 & 255).toString(16).padStart(2, "0")}${(A >> 16 & 255).toString(16).padStart(2, "0")}`
          };
        }, d = f(2), u = f(10), c = f(18), g = f(26);
        if (a + 46 + 4 <= n.length) {
          const p = i.getUint32(a + 42, !0);
          let y = a + 46;
          if (p & 1 && y + 4 <= n.length) {
            const m = i.getUint32(y, !0), A = m & 255, E = m >> 8 & 255, x = m >> 16 & 255;
            m !== 4294967295 && (l = `#${A.toString(16).padStart(2, "0")}${E.toString(16).padStart(2, "0")}${x.toString(16).padStart(2, "0")}`);
          }
        }
        this.borderFills.push({ left: d, right: u, top: c, bottom: g, faceColor: l });
      }
      a += h;
    }
  }
  parseBodyText() {
    const r = this.ole.FullPaths.filter(
      (n) => n.toLowerCase().includes("bodytext/section")
    ).sort((n, i) => {
      const a = n.toLowerCase().match(/section(\d+)/), s = i.toLowerCase().match(/section(\d+)/), o = a ? parseInt(a[1], 10) : 0, h = s ? parseInt(s[1], 10) : 0;
      return o - h;
    });
    if (r.length === 0)
      throw new Error("BodyText sections not found");
    let e = "";
    for (const n of r) {
      const i = lt.find(this.ole, n);
      if (!i || !i.content) continue;
      let a = new Uint8Array(i.content);
      if (this.compressed)
        try {
          a = Vt.inflate(a, { raw: !0 });
        } catch (S) {
          console.error(`Decompression failed for ${n}`, S);
          continue;
        }
      const s = new DataView(
        a.buffer,
        a.byteOffset,
        a.byteLength
      );
      let o = 0, h = 0, l = null, f = [], d = 0, u = !1, c = [], g = "", p = -1, y = -1, m = 1, A = 1, E = 0, x = -1;
      for (; o < a.length && !(o + 4 > a.length); ) {
        const S = s.getUint32(o, !0);
        o += 4;
        const R = S & 1023, v = S >>> 10 & 1023;
        let O = S >>> 20 & 4095;
        if (O === 4095) {
          if (o + 4 > a.length) break;
          O = s.getUint32(o, !0), o += 4;
        }
        if (O === 0)
          break;
        if (u && v < x) {
          if (l) {
            const C = f.length > 0 ? Wt(
              l,
              f,
              this.charShapes
            ) : zt(l);
            if (C) {
              let B = C;
              h === 2 ? B = `<div align="right">${B}</div>` : h === 3 && (B = `<div align="center">${B}</div>`), g += (g ? "<br>" : "") + B;
            }
            l = null, f = [];
          }
          if (p !== -1 && y !== -1) {
            for (; c.length <= p; ) c.push([]);
            for (; c[p].length <= y; )
              c[p].push(null);
            c[p][y] = {
              text: g.trim(),
              colSpan: m,
              rowSpan: A,
              borderFillId: E
            };
          }
          if (c.length > 0) {
            let C = 0;
            for (const B of c)
              B.length > C && (C = B.length);
            e += `
<table style="border-collapse: collapse; width: 100%;">
`;
            for (let B = 0; B < c.length; B++) {
              e += `  <tr>
`;
              for (let Z = 0; Z < C; Z++) {
                const T = c[B][Z];
                if (T) {
                  let _ = ` style="${this.buildCellStyle(T.borderFillId)}"`;
                  T.colSpan > 1 && (_ += ` colspan="${T.colSpan}"`), T.rowSpan > 1 && (_ += ` rowspan="${T.rowSpan}"`), e += `    <td${_}>${T.text}</td>
`;
                }
              }
              e += `  </tr>
`;
            }
            e += `</table>

`;
          }
          u = !1, c = [], g = "", p = -1, y = -1, E = 0, x = -1;
        }
        if (R === 66) {
          if (l && f.length > 0) {
            const C = Wt(
              l,
              f,
              this.charShapes
            );
            if (C) {
              let B = C;
              h === 2 ? B = `<div align="right">${B}</div>` : h === 3 && (B = `<div align="center">${B}</div>`), u ? g += (g ? "<br>" : "") + B : e += B + `

`;
            }
            l = null, f = [];
          } else if (l) {
            const C = zt(l);
            if (C) {
              let B = C;
              h === 2 ? B = `<div align="right">${B}</div>` : h === 3 && (B = `<div align="center">${B}</div>`), u ? g += (g ? "<br>" : "") + B : e += B + `

`;
            }
            l = null;
          }
          if (o + O <= a.length)
            try {
              const C = s.getUint32(o + 8, !0);
              C < this.paraShapes.length ? h = this.paraShapes[C].alignment : h = 0, o + 24 <= a.length ? d = s.getUint16(o + 22, !0) : d = 0, f = [];
            } catch {
            }
        } else if (R === 67)
          o + O <= a.length && (l = a.slice(o, o + O));
        else if (R === 68) {
          if (o + O <= a.length)
            for (let C = 0; C < d; C++) {
              const B = o + C * 8;
              if (B + 8 <= o + O && B + 8 <= a.length) {
                const Z = s.getUint32(B, !0), T = s.getUint32(B + 4, !0);
                f.push({ pos: Z, shapeId: T });
              }
            }
        } else if (R === 77) {
          if (l) {
            const C = f.length > 0 ? Wt(
              l,
              f,
              this.charShapes
            ) : zt(l);
            if (C) {
              let B = C;
              h === 2 ? B = `<div align="right">${B}</div>` : h === 3 && (B = `<div align="center">${B}</div>`), e += B + `

`;
            }
            l = null, f = [];
          }
          o + O <= a.length && (u = !0, c = [], g = "", p = -1, y = -1, E = 0, x = v);
        } else if (R === 72 && o + O <= a.length && u) {
          if (l) {
            const C = f.length > 0 ? Wt(
              l,
              f,
              this.charShapes
            ) : zt(l);
            if (C) {
              let B = C;
              h === 2 ? B = `<div align="right">${B}</div>` : h === 3 && (B = `<div align="center">${B}</div>`), g += (g ? "<br>" : "") + B;
            }
            l = null, f = [];
          }
          if (p !== -1 && y !== -1) {
            for (; c.length <= p; )
              c.push([]);
            for (; c[p].length <= y; )
              c[p].push(null);
            c[p][y] = {
              text: g.trim(),
              colSpan: m,
              rowSpan: A,
              borderFillId: E
            };
          }
          if (g = "", O >= 34) {
            const C = o + 8;
            y = s.getUint16(C, !0), p = s.getUint16(
              C + 2,
              !0
            ), m = s.getUint16(
              C + 4,
              !0
            ), A = s.getUint16(
              C + 6,
              !0
            ), E = s.getUint16(
              o + 32,
              !0
            );
          } else
            y = 0, p = 0, m = 1, A = 1, E = 0;
        }
        o += O;
      }
      if (l) {
        const S = f.length > 0 ? Wt(
          l,
          f,
          this.charShapes
        ) : zt(l);
        if (S) {
          let R = S;
          h === 2 ? R = `<div align="right">${R}</div>` : h === 3 && (R = `<div align="center">${R}</div>`), u ? g += (g ? "<br>" : "") + R : e += R + `

`;
        }
        l = null, f = [];
      }
      if (u) {
        if (p !== -1 && y !== -1) {
          for (; c.length <= p; ) c.push([]);
          for (; c[p].length <= y; )
            c[p].push(null);
          c[p][y] || (c[p][y] = {
            text: g.trim(),
            colSpan: m,
            rowSpan: A,
            borderFillId: E
          });
        }
        if (c.length > 0) {
          let S = 0;
          for (const R of c)
            R.length > S && (S = R.length);
          e += `
<table style="border-collapse: collapse; width: 100%;">
`;
          for (let R = 0; R < c.length; R++) {
            e += `  <tr>
`;
            for (let v = 0; v < S; v++) {
              const O = c[R][v];
              if (O) {
                let B = ` style="${this.buildCellStyle(O.borderFillId)}"`;
                O.colSpan > 1 && (B += ` colspan="${O.colSpan}"`), O.rowSpan > 1 && (B += ` rowspan="${O.rowSpan}"`), e += `    <td${B}>${O.text}</td>
`;
              }
            }
            e += `  </tr>
`;
          }
          e += `</table>

`;
        }
        u = !1, c = [], g = "", p = -1, y = -1, m = 1, A = 1, E = 0, x = -1;
      }
    }
    return e.trim();
  }
  // TODO: borderFillId를 사용하여 스타일을 생성해야함
  buildCellStyle(r) {
    const e = ["padding:8px"];
    if (r === void 0 || r < 0 || r >= this.borderFills.length)
      return e.push("border:1px solid #ddd"), e.join(";");
    const n = this.borderFills[r];
    if (!n)
      return e.push("border:1px solid #ddd"), e.join(";");
    const i = (a, s) => {
      if (s.type === 0 || s.thickness === 0) {
        e.push(`border-${a}:none`);
        return;
      }
      const o = Math.max(1, Math.round(s.thickness * 0.1) || 1);
      e.push(`border-${a}:${o}px solid ${s.color}`);
    };
    return i("left", n.left), i("right", n.right), i("top", n.top), i("bottom", n.bottom), n.faceColor && e.push(`background:${n.faceColor}`), e.join(";");
  }
};
async function Kl(t) {
  let r;
  t instanceof File || t instanceof Blob ? r = await t.arrayBuffer() : r = t;
  const e = new Uint8Array(r);
  try {
    return new Zl(e).parse();
  } catch (n) {
    throw console.error("HWP to MD conversion error parsing file", n), new Error(`Failed to convert HWP: ${n.message}`);
  }
}
const fh = () => {
  const [t, r] = Oe(!1), [e, n] = Oe(null), [i, a] = Oe(null);
  return { convert: async (o) => {
    r(!0), n(null), a(null);
    try {
      const h = await Kl(o);
      return a(h), h;
    } catch (h) {
      throw n(h), h;
    } finally {
      r(!1);
    }
  }, isLoading: t, error: e, result: i };
};
class Jl {
  constructor() {
    ce(this, "ole");
    ce(this, "compressed", !1);
    // 파싱된 정보 저장
    ce(this, "faceNames", []);
    ce(this, "charShapes", []);
    ce(this, "paraShapes", []);
    ce(this, "pageDefs", []);
    ce(this, "borderFills", []);
    ce(this, "imageRels", []);
  }
  /**
   * ArrayBuffer에서 HWP OLE를 읽고 DocInfo를 파싱합니다.
   * @returns 파싱된 OLE 컨테이너와 압축 여부
   */
  parse(r) {
    return this.ole = lt.read(r, { type: "buffer" }), this.parseFileHeader(), this.parseDocInfo(), this.extractImages(), {
      faceNames: this.faceNames,
      charShapes: this.charShapes,
      paraShapes: this.paraShapes,
      pageDefs: this.pageDefs,
      borderFills: this.borderFills,
      imageRels: this.imageRels,
      compressed: this.compressed,
      ole: this.ole
    };
  }
  /** OLE에서 BodyText 섹션 스트림 목록 반환 */
  getSectionStreams() {
    const r = this.ole.FullPaths.filter((n) => n.toLowerCase().includes("bodytext/section")).sort((n, i) => {
      const a = n.toLowerCase().match(/section(\d+)/), s = i.toLowerCase().match(/section(\d+)/), o = a ? parseInt(a[1], 10) : 0, h = s ? parseInt(s[1], 10) : 0;
      return o - h;
    });
    if (r.length === 0) throw new Error("BodyText sections not found");
    const e = [];
    for (const n of r) {
      const i = lt.find(this.ole, n);
      if (!i || !i.content) continue;
      let a = new Uint8Array(i.content);
      if (this.compressed)
        try {
          a = Vt.inflate(a, { raw: !0 });
        } catch {
          continue;
        }
      e.push(a);
    }
    return e;
  }
  // ─── FileHeader ──────────────────────────────────────────
  parseFileHeader() {
    const r = this.ole.FullPaths.find((s) => s.toLowerCase().endsWith("fileheader"));
    if (!r) throw new Error("Not a valid HWP 5.0 file");
    const e = lt.find(this.ole, r);
    if (!e || !e.content || e.content.length < 40)
      throw new Error("Invalid FileHeader");
    const n = new Uint8Array(e.content), a = new DataView(n.buffer, n.byteOffset, n.byteLength).getUint32(36, !0);
    this.compressed = (a & 1) !== 0;
  }
  // ─── DocInfo ──────────────────────────────────────────────
  parseDocInfo() {
    const r = this.ole.FullPaths.find((s) => s.toLowerCase().endsWith("docinfo"));
    if (!r) return;
    const e = lt.find(this.ole, r);
    if (!e || !e.content) return;
    let n = new Uint8Array(e.content);
    if (this.compressed)
      try {
        n = Vt.inflate(n, { raw: !0 });
      } catch {
      }
    const i = new DataView(n.buffer, n.byteOffset, n.byteLength);
    let a = 0;
    for (; a < n.length && !(a + 4 > n.length); ) {
      const s = i.getUint32(a, !0);
      a += 4;
      const o = s & 1023;
      let h = s >>> 20 & 4095;
      if (h === 4095) {
        if (a + 4 > n.length) break;
        h = i.getUint32(a, !0), a += 4;
      }
      if (h !== 0) {
        if (a + h > n.length) break;
        o === 17 ? this.parseFaceNameRecord(i, a, h) : o === 21 ? this.parseCharShapeRecord(i, a, h) : o === 25 ? this.parseParaShapeRecord(i, a, h) : o === 20 && this.parseBorderFillRecord(i, a, h, n.length), a += h;
      }
    }
  }
  // ─── 개별 레코드 파싱 ─────────────────────────────────────
  /** HWPTAG_FACE_NAME(17) - 글꼴 이름 추출 */
  parseFaceNameRecord(r, e, n) {
    try {
      if (n < 3) return;
      const i = r.getUint16(e + 1, !0);
      if (i === 0 || e + 3 + i * 2 > e + n) return;
      const a = new Uint8Array(r.buffer, r.byteOffset + e + 3, i * 2), s = new TextDecoder("utf-16le").decode(a).replace(/\0/g, "");
      this.faceNames.push(s);
    } catch {
    }
  }
  /**
   * HWPTAG_CHAR_SHAPE(21) - 글자 모양 파싱
   * 오프셋 레이아웃:
   *   0~13:  WORD[7]  언어별 글꼴 ID (14바이트)
   *   14~20: UINT8[7] 장평 비율
   *   21~27: INT8[7]  자간
   *   28~34: UINT8[7] 상대크기
   *   35~41: INT8[7]  글자 위치
   *   42~45: INT32    기준 크기 (HWPUNIT, /100 → pt)
   *   46~49: UINT32   속성 (굵기, 기울임, 밑줄, 취소선 등)
   *   50~51: INT8×2   그림자 간격
   *   52~55: COLORREF 글자 색
   */
  parseCharShapeRecord(r, e, n) {
    try {
      const i = [];
      for (let u = 0; u < 7; u++)
        e + u * 2 + 2 <= e + n ? i.push(r.getUint16(e + u * 2, !0)) : i.push(0);
      const a = n >= 22 ? r.getInt8(e + 21) : 0;
      let s = 10;
      if (n >= 46) {
        const u = r.getInt32(e + 42, !0);
        s = Math.round(u / 100);
      }
      let o = !1, h = !1, l = !1, f = !1;
      if (n >= 50) {
        const u = r.getUint32(e + 46, !0);
        h = (u & 1) !== 0, o = (u & 2) !== 0, l = (u & 12) !== 0, f = (u & 7 << 18) !== 0;
      }
      let d = "#000000";
      if (n >= 56) {
        const u = r.getUint32(e + 52, !0);
        d = `#${(u & 255).toString(16).padStart(2, "0")}${(u >> 8 & 255).toString(16).padStart(2, "0")}${(u >> 16 & 255).toString(16).padStart(2, "0")}`;
      }
      this.charShapes.push({ fontIds: i, isBold: o, isItalic: h, isUnderline: l, isStrike: f, baseSize: s, charColor: d, letterSpacing: a });
    } catch {
    }
  }
  /**
   * HWPTAG_PARA_SHAPE(25) - 문단 모양 파싱
   * 레이아웃:
   *   +0  UINT32 속성1 (bit0-1: 줄간격종류, bit2-4: 정렬)
   *   +4  INT32  왼쪽 여백
   *   +8  INT32  오른쪽 여백
   *   +12 INT32  들여쓰기/내어쓰기
   *   +16 INT32  문단 간격 위
   *   +20 INT32  문단 간격 아래
   *   +24 INT32  줄 간격 (5.0.2.5 미만)
   *   +28 UINT16 탭 정의 ID
   *   +30 UINT16 번호/글머리표 ID
   *   +32 UINT16 테두리/배경 ID
   *   +34 INT16  테두리 좌 간격
   *   +36 INT16  테두리 우 간격
   *   +38 INT16  테두리 상 간격
   *   +40 INT16  테두리 하 간격
   *   +42 UINT32 속성2 (5.0.1.7+)
   *   +46 UINT32 속성3 (5.0.2.5+)
   *   +50 UINT32 줄 간격 (5.0.2.5+)
   */
  parseParaShapeRecord(r, e, n) {
    try {
      const i = r.getUint32(e, !0), a = i & 3, s = i >>> 2 & 7, o = n >= 8 ? r.getInt32(e + 4, !0) : 0, h = n >= 12 ? r.getInt32(e + 8, !0) : 0, l = n >= 16 ? r.getInt32(e + 12, !0) : 0, f = n >= 20 ? r.getInt32(e + 16, !0) : 0, d = n >= 24 ? r.getInt32(e + 20, !0) : 0, u = n >= 28 ? r.getInt32(e + 24, !0) : 0, c = n >= 34 ? r.getUint16(e + 32, !0) : 0, g = n >= 36 ? r.getInt16(e + 34, !0) : 0, p = n >= 38 ? r.getInt16(e + 36, !0) : 0, y = n >= 40 ? r.getInt16(e + 38, !0) : 0, m = n >= 42 ? r.getInt16(e + 40, !0) : 0, A = n >= 54 ? r.getUint32(e + 50, !0) : 0, E = A > 0 ? A : u;
      this.paraShapes.push({
        alignment: s,
        leftMargin: o,
        rightMargin: h,
        indent: l,
        lineHeight: E,
        lineHeightType: a,
        spaceAbove: f,
        spaceBelow: d,
        borderFillId: c,
        borderDistLeft: g,
        borderDistRight: p,
        borderDistTop: y,
        borderDistBottom: m
      });
    } catch {
    }
  }
  /** HWPTAG_BORDER_FILL(20) - 테두리/배경 */
  parseBorderFillRecord(r, e, n, i) {
    try {
      const a = (d) => {
        if (e + d + 8 > i) return { type: 0, thickness: 0, color: "#000000" };
        const u = r.getUint16(e + d, !0), c = r.getUint16(e + d + 2, !0), g = r.getUint32(e + d + 4, !0);
        return {
          type: u,
          thickness: c,
          color: `#${(g & 255).toString(16).padStart(2, "0")}${(g >> 8 & 255).toString(16).padStart(2, "0")}${(g >> 16 & 255).toString(16).padStart(2, "0")}`
        };
      };
      let s;
      const o = a(2), h = a(10), l = a(18), f = a(26);
      if (n >= 50) {
        const d = r.getUint32(e + 42, !0), u = e + 46;
        if (d & 1 && u + 4 <= i) {
          const c = r.getUint32(u, !0);
          if (c !== 4294967295) {
            const g = c & 255, p = c >> 8 & 255, y = c >> 16 & 255;
            s = `#${g.toString(16).padStart(2, "0")}${p.toString(16).padStart(2, "0")}${y.toString(16).padStart(2, "0")}`;
          }
        }
      }
      this.borderFills.push({ left: o, right: h, top: l, bottom: f, faceColor: s });
    } catch {
    }
  }
  // ─── 이미지 추출 ──────────────────────────────────────────
  /** OLE BinData 스토리지에서 이미지 파일 추출 */
  extractImages() {
    const r = this.ole.FullPaths.filter((e) => e.toLowerCase().includes("bindata/"));
    for (let e = 0; e < r.length; e++) {
      const n = r[e], i = lt.find(this.ole, n);
      if (!i || !i.content) continue;
      let a = new Uint8Array(i.content);
      try {
        a[0] === 120 && (a = Vt.inflate(a));
      } catch (f) {
        console.warn(`이미지 압축 해제 실패 (${n}):`, f);
      }
      const s = n.match(/\.([a-zA-Z0-9]+)$/), o = s ? s[1].toLowerCase() : "png", h = n.match(/BIN(\d+)/i), l = h ? parseInt(h[1], 10) : e + 1;
      this.imageRels.push({
        id: `rIdImg${l}`,
        target: `media/image${l}.${o}`,
        data: a,
        ext: o
      });
    }
  }
}
class Vl {
  constructor(r, e, n, i, a = []) {
    ce(this, "faceNames");
    ce(this, "charShapes");
    ce(this, "paraShapes");
    ce(this, "pageDefs");
    ce(this, "borderFills");
    ce(this, "nextImageId", 1);
    this.faceNames = r, this.charShapes = e, this.paraShapes = n, this.pageDefs = i, this.borderFills = a;
  }
  // ─── 섹션 파싱 + XML 생성 ──────────────────────────────────
  /** 섹션 하나의 바이너리 스트림을 파싱하여 DOCX body/header/footer XML을 반환 */
  parseSection(r) {
    const e = new DataView(r.buffer, r.byteOffset, r.byteLength);
    let n = 0, i = "", a = 0, s = null, o = [], h = [], l = !1, f = [], d = [], u = -1, c = -1, g = 1, p = 1, y = 0, m = 0, A = 0, E = -1, x = 0, S = 0, R = null, v = -1, O = "", C = "", B = 0, Z = 0;
    const T = () => {
      if (!s) return "";
      const se = this.buildParagraphXml(s, o, a, l, h);
      return s = null, o = [], h = [], se;
    }, j = (se) => {
      if (se) {
        if (l) {
          d.push(se);
          return;
        }
        if (R === "header") {
          O += se;
          return;
        }
        if (R === "footer") {
          C += se;
          return;
        }
        i += se;
      }
    }, _ = () => {
      if (f.length === 0) return;
      const se = this.buildTableXml(f, x, S);
      f = [], S = 0, R === "header" ? O += se : R === "footer" ? C += se : i += se;
    };
    for (; n < r.length && !(n + 4 > r.length); ) {
      const se = e.getUint32(n, !0);
      n += 4;
      const P = se & 1023, G = se >>> 10 & 1023;
      let U = se >>> 20 & 4095;
      if (U === 4095) {
        if (n + 4 > r.length) break;
        U = e.getUint32(n, !0), n += 4;
      }
      if (U !== 0) {
        if (n + U > r.length) break;
        if (R !== null && G <= v)
          j(T()), R = null, v = -1;
        else if (l && G < E) {
          if (u !== -1 && c !== -1) {
            const Q = T();
            Q && d.push(Q), this.fillTableData(f, u, c, {
              text: d.join(""),
              colSpan: g,
              rowSpan: p,
              borderFillId: y,
              width: m,
              height: A
            });
          }
          _(), l = !1, d = [], u = -1, c = -1, E = -1;
        } else if (P === 71) {
          if (U >= 4) {
            const Q = e.getUint32(n, !0);
            Q === 1684104552 ? (R = "header", v = G, O = "") : Q === 1953460070 && (R = "footer", v = G, C = "");
          }
        } else if (P === 73)
          U >= 40 && this.pageDefs.push({
            width: Math.round(e.getUint32(n, !0) / 5),
            height: Math.round(e.getUint32(n + 4, !0) / 5),
            leftMargin: Math.round(e.getUint32(n + 8, !0) / 5),
            rightMargin: Math.round(e.getUint32(n + 12, !0) / 5),
            topMargin: Math.round(e.getUint32(n + 16, !0) / 5),
            bottomMargin: Math.round(e.getUint32(n + 20, !0) / 5),
            headerMargin: Math.round(e.getUint32(n + 24, !0) / 5),
            footerMargin: Math.round(e.getUint32(n + 28, !0) / 5)
          });
        else if (P === 76) {
          if (U >= 26)
            try {
              const Q = e.getInt32(n + 18, !0), M = e.getInt32(n + 22, !0);
              Q > 0 && M > 0 && (B = Math.round(Q / 5), Z = Math.round(M / 5));
            } catch (Q) {
              console.error("HWPTAG_SHAPE_COMPONENT 파싱 에러:", Q);
            }
        } else if (P === 85) {
          if (U >= 73)
            try {
              const Q = e.getUint16(n + 71, !0), M = Math.abs(e.getInt32(n + 20, !0) - e.getInt32(n + 12, !0)), X = Math.abs(e.getInt32(n + 40, !0) - e.getInt32(n + 16, !0)), de = this.pageDefs.length > 0 ? this.pageDefs[this.pageDefs.length - 1] : null, le = de ? Math.max(1, de.width - de.leftMargin - de.rightMargin) : 8306, ne = Math.round(le / 2), ye = Math.round(ne * 3 / 4);
              let Te = M > 0 ? Math.round(M / 5) : B > 0 ? B : ne, _e = X > 0 ? Math.round(X / 5) : Z > 0 ? Z : ye;
              if (Te > le) {
                const Se = le / Te;
                Te = le, _e = Math.round(_e * Se);
              }
              j(this.buildImageXml(Q, Te, _e));
            } catch (Q) {
              console.error("HWPTAG_SHAPE_COMPONENT_PICTURE 파싱 에러:", Q);
            }
          B = 0, Z = 0;
        } else if (P === 66) {
          if (j(T()), U >= 12)
            try {
              a = e.getUint32(n + 8, !0);
            } catch {
            }
        } else if (P === 67)
          s = r.slice(n, n + U);
        else if (P === 68) {
          o = [];
          const Q = Math.floor(U / 8);
          for (let M = 0; M < Q; M++) {
            const X = e.getUint32(n + M * 8, !0), de = e.getUint32(n + M * 8 + 4, !0);
            o.push({ pos: X, shapeId: de });
          }
        } else if (P === 69) {
          h = [];
          const Q = 36, M = Math.floor(U / Q);
          for (let X = 0; X < M; X++) {
            const de = n + X * Q;
            if (de + Q > n + U) break;
            h.push({
              textOffset: e.getInt32(de, !0),
              lineY: e.getInt32(de + 4, !0),
              lineHeight: e.getInt32(de + 8, !0),
              textHeight: e.getInt32(de + 12, !0),
              baselineOffset: e.getInt32(de + 16, !0),
              lineSpacing: e.getInt32(de + 20, !0),
              columnStart: e.getInt32(de + 24, !0),
              segmentWidth: e.getInt32(de + 28, !0),
              flags: e.getUint32(de + 32, !0)
            });
          }
        } else if (P === 77) {
          j(T());
          let Q = !1;
          if (U >= 4)
            try {
              Q = (e.getUint32(n, !0) & 1) !== 0;
            } catch {
            }
          if (U >= 22)
            try {
              S = e.getUint16(n + 20, !0);
            } catch {
            }
          x = Q && a < this.paraShapes.length ? this.paraShapes[a].alignment : 0, l = !0, f = [], d = [], u = -1, c = -1, y = 0, m = 0, A = 0, E = G;
        } else if (P === 72 && l) {
          const Q = T();
          if (Q && d.push(Q), u !== -1 && c !== -1 && this.fillTableData(f, u, c, {
            text: d.join(""),
            colSpan: g,
            rowSpan: p,
            borderFillId: y,
            width: m,
            height: A
          }), d = [], U >= 26)
            try {
              const M = n + 8;
              c = e.getUint16(M, !0), u = e.getUint16(M + 2, !0), g = e.getUint16(M + 4, !0), p = e.getUint16(M + 6, !0), m = e.getUint32(M + 8, !0), A = e.getUint32(M + 12, !0), U >= 34 && (y = e.getUint16(M + 24, !0));
            } catch {
            }
        }
        n += U;
      }
    }
    const K = T();
    return l ? (K && d.push(K), u !== -1 && c !== -1 && this.fillTableData(f, u, c, {
      text: d.join(""),
      colSpan: g,
      rowSpan: p,
      borderFillId: y,
      width: m,
      height: A
    }), _()) : j(K), {
      body: i,
      header: O || null,
      footer: C || null
    };
  }
  // ─── 문단 XML 빌드 ──────────────────────────────────────────
  /** PARA_TEXT + PARA_CHAR_SHAPE + PARA_SHAPE + LINE_SEG → 완전한 <w:p> */
  buildParagraphXml(r, e, n, i = !1, a = []) {
    const h = a.length > 0 && (a[0].flags & 16) !== 0, l = /* @__PURE__ */ new Map();
    for (let u = 1; u < a.length; u++) {
      const c = a[u];
      c.flags & 16 ? l.set(c.textOffset, "page") : c.flags & 32 && l.set(c.textOffset, "column");
    }
    const f = this.buildRunsFromText(r, e, l);
    if (!f) return "";
    let d = "<w:p>";
    return d += this.buildParaPropsXml(n, i, h), d += f, d += "</w:p>", d;
  }
  /** ParaShape → <w:pPr> */
  buildParaPropsXml(r, e = !1, n = !1) {
    if (r < 0 || r >= this.paraShapes.length) return "";
    const i = this.paraShapes[r];
    let a = "";
    n && (a += "<w:pageBreakBefore/>");
    const s = this.alignToDocx(i.alignment);
    s !== "both" && (a += `<w:jc w:val="${s}"/>`);
    {
      const o = i.indent > 0 ? Math.round(i.indent / 5) : 0, h = i.indent < 0 ? Math.abs(Math.round(i.indent / 5)) : 0;
      let l, f;
      if (e ? (l = h > 0 ? h : 0, f = 0) : (l = Math.max(h, Math.round(i.leftMargin / 5)), f = Math.max(0, Math.round(i.rightMargin / 5))), l !== void 0 || f !== void 0 || o !== void 0 || h !== void 0) {
        let d = "";
        l !== void 0 && (d += ` w:left="${l}"`), f !== void 0 && (d += ` w:right="${f}"`), h !== void 0 ? d += ` w:hanging="${h}"` : o !== void 0 && (d += ` w:firstLine="${o}"`), a += `<w:ind${d}/>`;
      }
    }
    {
      const o = i.spaceAbove ? Math.max(0, Math.round(i.spaceAbove / 5)) : 0, h = i.spaceBelow ? Math.max(0, Math.round(i.spaceBelow / 5)) : 0;
      let l = 0, f = "";
      if (i.lineHeight > 0 && (i.lineHeightType === 0 ? (l = Math.round(i.lineHeight * 240 / 100), f = "auto") : i.lineHeightType === 1 ? (l = Math.round(i.lineHeight / 5), f = "exact") : i.lineHeightType === 2 && (l = Math.round(i.lineHeight / 5), f = "atLeast")), o || h || l) {
        let d = "";
        o && (d += ` w:before="${o}"`), h && (d += ` w:after="${h}"`), l && (d += ` w:line="${l}" w:lineRule="${f}"`), a += `<w:spacing${d}/>`;
      }
    }
    if (i.borderFillId > 0 && i.borderFillId <= this.borderFills.length) {
      const o = this.borderFills[i.borderFillId - 1];
      o.faceColor && (a += `<w:shd w:val="clear" w:color="auto" w:fill="${o.faceColor.replace("#", "")}"/>`);
      const h = [
        { key: "top", border: o.top },
        { key: "left", border: o.left },
        { key: "bottom", border: o.bottom },
        { key: "right", border: o.right }
      ];
      if (h.some((f) => f.border && f.border.type !== 0)) {
        a += "<w:pBdr>";
        for (const { key: f, border: d } of h)
          d && d.type !== 0 && (a += this.buildBorderXml(f, d));
        a += "</w:pBdr>";
      }
    }
    return a ? `<w:pPr>${a}</w:pPr>` : "";
  }
  // ─── Run XML 빌드 ──────────────────────────────────────────
  /** PARA_TEXT 바이너리에서 <w:r> 요소들 직접 생성
   * @param inlineBreaks  LINE_SEG에서 추출한 {문자위치 → 'page'|'column'} 맵.
   *                      해당 위치 직전에 <w:br w:type="..."/> 를 삽입한다.
   */
  buildRunsFromText(r, e, n = /* @__PURE__ */ new Map()) {
    const i = new DataView(r.buffer, r.byteOffset, r.byteLength), a = Math.floor(r.length / 2);
    if (a === 0) return "";
    let s = "", o = "", h = e.length > 0 ? e[0].shapeId : -1, l = e.length > 0 && e[0].pos === 0 ? 1 : 0;
    const f = () => {
      o && (s += this.buildRunXml(o, h), o = "");
    };
    let d = 0;
    for (; d < a; ) {
      const u = n.get(d);
      u && (f(), s += `<w:r><w:br w:type="${u}"/></w:r>`), l < e.length && e[l].pos === d && (f(), h = e[l].shapeId, l++);
      const c = i.getUint16(d * 2, !0);
      c === 10 || c === 13 ? (f(), d += 1) : c >= 0 && c <= 31 ? [0, 10, 13, 24, 25, 26, 27, 28, 29, 30, 31].includes(c) ? d += 1 : d += 8 : (o += String.fromCharCode(c), d += 1);
    }
    return f(), s;
  }
  /** 텍스트 + CharShape ID → <w:r> XML (서식 포함) */
  buildRunXml(r, e) {
    let n = "<w:r>";
    if (e >= 0 && e < this.charShapes.length) {
      const i = this.charShapes[e];
      let a = "";
      if (i.fontIds && i.fontIds.length > 0) {
        const s = i.fontIds[0], o = i.fontIds[1], h = s < this.faceNames.length ? this.faceNames[s] : "", l = o < this.faceNames.length ? this.faceNames[o] : "";
        if (h || l) {
          const f = h || l, d = l || h;
          a += `<w:rFonts w:ascii="${this.escapeXml(d)}" w:eastAsia="${this.escapeXml(f)}" w:hAnsi="${this.escapeXml(d)}"/>`;
        }
      }
      if (i.isBold && (a += "<w:b/><w:bCs/>"), i.isItalic && (a += "<w:i/><w:iCs/>"), i.isUnderline && (a += '<w:u w:val="single"/>'), i.isStrike && (a += "<w:strike/>"), i.baseSize > 0) {
        const s = i.baseSize * 2;
        a += `<w:sz w:val="${s}"/>`, a += `<w:szCs w:val="${s}"/>`;
      }
      if (i.charColor && i.charColor !== "#000000" && (a += `<w:color w:val="${i.charColor.replace("#", "")}"/>`), i.letterSpacing && i.letterSpacing !== 0) {
        const s = Math.round(i.letterSpacing);
        a += `<w:spacing w:val="${s}"/>`;
      }
      a && (n += `<w:rPr>${a}</w:rPr>`);
    }
    return n += `<w:t xml:space="preserve">${this.escapeXml(r)}</w:t>`, n += "</w:r>", n;
  }
  // ─── 테이블 XML 빌드 ────────────────────────────────────────
  /** tableData 2D 배열 → 완전한 <w:tbl> XML */
  buildTableXml(r, e = 0, n = 0) {
    let i = 0;
    for (const f of r)
      f.length > i && (i = f.length);
    if (i === 0) return "";
    const a = new Array(i).fill(0);
    for (const f of r) {
      let d = 0;
      for (; d < i; ) {
        const u = d < f.length ? f[d] : null;
        if (u && !u.isVMergeContinue) {
          const c = Math.max(1, u.colSpan || 1);
          if (c === 1 && u.width > 0) {
            const g = Math.round(u.width / 5);
            g > a[d] && (a[d] = g);
          }
          d += c;
        } else
          d++;
      }
    }
    for (const f of r) {
      let d = 0;
      for (; d < i; ) {
        const u = d < f.length ? f[d] : null;
        if (u && !u.isVMergeContinue) {
          const c = Math.min(Math.max(1, u.colSpan || 1), i - d);
          if (c > 1 && u.width > 0) {
            const g = Math.round(u.width / 5);
            let p = 0, y = 0;
            for (let m = 0; m < c; m++)
              a[d + m] === 0 ? p++ : y += a[d + m];
            if (p > 0) {
              const m = Math.max(1, Math.round((g - y) / p));
              for (let A = 0; A < c; A++)
                a[d + A] === 0 && (a[d + A] = m);
            }
          }
          d += c;
        } else
          d++;
      }
    }
    for (let f = 0; f < i; f++)
      a[f] === 0 && (a[f] = Math.round(9e3 / i));
    const s = a.reduce((f, d) => f + d, 0), o = new Array(r.length).fill(0);
    for (let f = 0; f < r.length; f++) {
      const d = r[f];
      for (const u of d)
        if (u && !u.isVMergeContinue && (u.rowSpan || 1) === 1 && u.height > 0) {
          const c = Math.round(u.height / 5);
          c > o[f] && (o[f] = c);
        }
      if (o[f] === 0) {
        for (const u of d)
          if (u && !u.isVMergeContinue && u.height > 0) {
            o[f] = Math.round(u.height / Math.max(1, u.rowSpan || 1) / 5);
            break;
          }
      }
    }
    let h = "<w:tbl>";
    h += "<w:tblPr>";
    const l = this.alignToDocx(e);
    if (l !== "both" && (h += `<w:jc w:val="${l}"/>`), h += `<w:tblW w:w="${s}" w:type="dxa"/><w:tblBorders>`, n > 0 && n <= this.borderFills.length) {
      const f = this.borderFills[n - 1];
      h += this.buildBorderXml("top", f.top), h += this.buildBorderXml("left", f.left), h += this.buildBorderXml("bottom", f.bottom), h += this.buildBorderXml("right", f.right), h += this.buildBorderXml("insideH", f.top), h += this.buildBorderXml("insideV", f.left);
    } else
      h += '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>', h += '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>', h += '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>', h += '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>', h += '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>', h += '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
    h += '</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>', h += "<w:tblGrid>";
    for (let f = 0; f < i; f++)
      h += `<w:gridCol w:w="${a[f]}"/>`;
    h += "</w:tblGrid>";
    for (let f = 0; f < r.length; f++) {
      h += "<w:tr>", o[f] > 0 && (h += `<w:trPr><w:trHeight w:val="${o[f]}" w:hRule="atLeast"/></w:trPr>`);
      for (let d = 0; d < i; d++) {
        const u = r[f][d];
        if (!u) continue;
        h += "<w:tc><w:tcPr>";
        const c = Math.min(Math.max(1, u.colSpan || 1), i - d);
        let g = 0;
        for (let p = 0; p < c; p++)
          g += a[d + p] || 0;
        if (h += `<w:tcW w:w="${g}" w:type="dxa"/>`, (u.colSpan || 1) > 1 && (h += `<w:gridSpan w:val="${u.colSpan}"/>`), (u.rowSpan || 1) > 1 && !u.isVMergeContinue ? h += '<w:vMerge w:val="restart"/>' : u.isVMergeContinue && (h += "<w:vMerge/>"), u.borderFillId > 0 && u.borderFillId <= this.borderFills.length) {
          const p = this.borderFills[u.borderFillId - 1];
          h += "<w:tcBorders>", h += this.buildBorderXml("top", p.top), h += this.buildBorderXml("left", p.left), h += this.buildBorderXml("bottom", p.bottom), h += this.buildBorderXml("right", p.right), h += "</w:tcBorders>", p.faceColor && (h += `<w:shd w:val="clear" w:color="auto" w:fill="${p.faceColor.replace("#", "")}"/>`);
        }
        if (h += "</w:tcPr>", u.isVMergeContinue)
          h += "<w:p/>";
        else {
          const p = u.text;
          h += p || "<w:p/>";
        }
        h += "</w:tc>", (u.colSpan || 1) > 1 && (d += u.colSpan - 1);
      }
      h += "</w:tr>";
    }
    return h += "</w:tbl>", h;
  }
  /** 2D 배열에 셀 데이터 채우기 (병합 처리 포함) */
  fillTableData(r, e, n, i) {
    for (; r.length <= e + i.rowSpan - 1; ) r.push([]);
    for (; r[e].length <= n; ) r[e].push(null);
    if (r[e][n] = i, i.rowSpan > 1)
      for (let a = 1; a < i.rowSpan; a++) {
        for (; r[e + a].length <= n; ) r[e + a].push(null);
        r[e + a][n] = { ...i, text: "", isVMergeContinue: !0 };
      }
    if (i.colSpan > 1)
      for (let a = 0; a < i.rowSpan; a++)
        for (let s = 1; s < i.colSpan; s++)
          for (; r[e + a].length <= n + s; ) r[e + a].push(null);
  }
  // ─── 이미지 XML ──────────────────────────────────────────
  /** DOCX DrawingML 이미지 요소 생성 */
  buildImageXml(r, e, n) {
    const i = this.nextImageId++, a = `rIdImg${r}`, s = Math.round(e * 635), o = Math.round(n * 635);
    return `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${s}" cy="${o}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${i}" name="Picture ${i}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${i}" name="Image"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${a}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${s}" cy="${o}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
  }
  // ─── 구역/페이지 설정 ────────────────────────────────────────
  /** 구역 설정 <w:sectPr> XML */
  buildSectPrXml(r = !1, e = !1) {
    const n = this.pageDefs.length > 0 ? this.pageDefs[this.pageDefs.length - 1] : null;
    let i = "<w:sectPr>";
    return r && (i += '<w:headerReference w:type="default" r:id="rIdHdr1"/>'), e && (i += '<w:footerReference w:type="default" r:id="rIdFtr1"/>'), n ? (i += `<w:pgSz w:w="${n.width}" w:h="${n.height}"/>`, i += `<w:pgMar w:top="${n.topMargin}" w:right="${n.rightMargin}" w:bottom="${n.bottomMargin}" w:left="${n.leftMargin}" w:header="${n.headerMargin}" w:footer="${n.footerMargin}" w:gutter="0"/>`) : (i += '<w:pgSz w:w="11906" w:h="16838"/>', i += '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="851" w:gutter="0"/>'), i += "</w:sectPr>", i;
  }
  // ─── XML 유틸리티 ────────────────────────────────────────────
  escapeXml(r) {
    return r.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
  }
  alignToDocx(r) {
    switch (r) {
      case 1:
        return "left";
      case 2:
        return "right";
      case 3:
        return "center";
      case 4:
        return "distribute";
      case 5:
        return "thaiDistribute";
      default:
        return "both";
    }
  }
  hwpBorderTypeToDocx(r) {
    switch (r) {
      case 0:
        return "nil";
      case 1:
        return "single";
      case 2:
        return "dashed";
      case 3:
        return "dotted";
      case 4:
        return "dotDash";
      case 5:
        return "dotDotDash";
      case 6:
        return "double";
      case 7:
        return "thick";
      default:
        return "single";
    }
  }
  /** HWP 선 굵기 코드 → DOCX w:sz (1/8pt) */
  hwpThicknessToDocxSz(r) {
    return [2, 3, 3, 5, 6, 7, 9, 11, 14, 16, 23, 34, 45, 68, 91, 113][r] ?? 4;
  }
  /** BorderFill의 한 방향 → DOCX <w:${side}> XML 문자열 */
  buildBorderXml(r, e) {
    if (!e || e.type === 0)
      return `<w:${r} w:val="nil"/>`;
    const n = this.hwpBorderTypeToDocx(e.type), i = this.hwpThicknessToDocxSz(e.thickness), a = e.color.replace("#", "");
    return `<w:${r} w:val="${n}" w:sz="${i}" w:space="0" w:color="${a}"/>`;
  }
}
const vr = class vr {
  async convert(r) {
    let e;
    r instanceof File || r instanceof Blob ? e = await r.arrayBuffer() : e = r;
    const n = new Uint8Array(e), i = new Jl(), a = i.parse(n), s = new Vl(
      a.faceNames,
      a.charShapes,
      a.paraShapes,
      a.pageDefs,
      a.borderFills
    ), o = i.getSectionStreams();
    let h = "", l = null, f = null;
    for (const d of o) {
      const u = s.parseSection(d);
      h += u.body, u.header !== null && (l = u.header), u.footer !== null && (f = u.footer);
    }
    return h += s.buildSectPrXml(l !== null, f !== null), this.createDocxPackage(h, a.imageRels, l, f, a.faceNames);
  }
  // ─── DOCX 패키징 ──────────────────────────────────────────
  async createDocxPackage(r, e, n, i, a = []) {
    const s = new yt(), o = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', h = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    s.file("[Content_Types].xml", this.contentTypesXml(e, n !== null, i !== null)), s.file("_rels/.rels", this.topRelsXml()), s.file("word/document.xml", this.documentXml(r)), s.file("word/_rels/document.xml.rels", this.documentRelsXml(e, n !== null, i !== null)), s.file("word/styles.xml", this.stylesXml(a)), s.file(
      "word/settings.xml",
      `${o}<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/></w:settings>`
    ), s.file("word/fontTable.xml", this.fontTableXml(a)), n !== null && s.file("word/header1.xml", `${o}<w:hdr ${h}>${n}</w:hdr>`), i !== null && s.file("word/footer1.xml", `${o}<w:ftr ${h}>${i}</w:ftr>`);
    for (const l of e)
      s.file(`word/${l.target}`, l.data);
    return await s.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });
  }
  // ─── 한국어 폰트 판별 ─────────────────────────────────────
  isKoreanFont(r) {
    return /함초롬|맑은|나눔|굴림|바탕|돋움|궁서|HCR|새굴림|윤|한컴|[가-힣]/.test(r);
  }
  // ─── 기본 폰트 결정 (문서 faceNames 기반) ─────────────────
  getDefaultFonts(r) {
    const e = r.filter((a) => a && a.trim()), n = e.find((a) => this.isKoreanFont(a)) || "Malgun Gothic", i = e.find((a) => !this.isKoreanFont(a) && /^[A-Za-z]/.test(a)) || n;
    return { korean: n, latin: i };
  }
  contentTypesXml(r, e = !1, n = !1) {
    const i = /* @__PURE__ */ new Set();
    for (const s of r)
      s.ext && i.add(s.ext.toLowerCase());
    let a = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>';
    for (const s of i) {
      const o = vr.MIME_MAP[s] ?? `image/${s}`;
      a += `<Default Extension="${s}" ContentType="${o}"/>`;
    }
    return a += '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>', e && (a += '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'), n && (a += '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'), a += "</Types>", a;
  }
  topRelsXml() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
  }
  documentXml(r) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${r}</w:body></w:document>`;
  }
  documentRelsXml(r, e = !1, n = !1) {
    let i = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    i += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>', i += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>', i += '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>', e && (i += '<Relationship Id="rIdHdr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'), n && (i += '<Relationship Id="rIdFtr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>');
    for (const a of r)
      i += `<Relationship Id="${a.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${a.target}"/>`;
    return i += "</Relationships>", i;
  }
  /** styles.xml: 문서의 실제 폰트를 기본값으로 사용 */
  stylesXml(r = []) {
    const { korean: e, latin: n } = this.getDefaultFonts(r);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${n}" w:eastAsia="${e}" w:hAnsi="${n}" w:cs="${e}"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults></w:styles>`;
  }
  /** fontTable.xml: 문서에 사용된 폰트 전체 목록 등록 */
  fontTableXml(r = []) {
    const e = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', n = [...new Set(r.filter((a) => a && a.trim()))];
    n.length === 0 && n.push("Malgun Gothic");
    let i = `${e}<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`;
    for (const a of n) {
      const s = this.isKoreanFont(a) ? "81" : "00";
      i += `<w:font w:name="${a}"><w:charset w:val="${s}"/><w:family w:val="auto"/><w:pitch w:val="variable"/></w:font>`;
    }
    return i += "</w:fonts>", i;
  }
};
// ─── MIME 타입 매핑 ───────────────────────────────────────
ce(vr, "MIME_MAP", {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  bmp: "image/bmp",
  gif: "image/gif",
  tiff: "image/tiff",
  tif: "image/tiff",
  emf: "image/x-emf",
  wmf: "image/x-wmf",
  svg: "image/svg+xml",
  webp: "image/webp"
});
let cn = vr;
async function Yl(t) {
  return new cn().convert(t);
}
const dh = () => {
  const [t, r] = Oe(!1), [e, n] = Oe(null), [i, a] = Oe(null);
  return { convert: async (o) => {
    r(!0), n(null), a(null);
    try {
      const h = await Yl(o);
      return a(h), h;
    } catch (h) {
      throw n(h), h;
    } finally {
      r(!1);
    }
  }, isLoading: t, error: e, result: i };
};
class ql {
  constructor() {
    ce(this, "ns", {
      hp: "http://www.hancom.co.kr/hwpml/2011/paragraph",
      hh: "http://www.hancom.co.kr/hwpml/2011/head",
      hc: "http://www.hancom.co.kr/hwpml/2011/core"
    });
    ce(this, "charProperties", /* @__PURE__ */ new Map());
    ce(this, "paraProperties", /* @__PURE__ */ new Map());
    ce(this, "styles", /* @__PURE__ */ new Map());
    ce(this, "images", /* @__PURE__ */ new Map());
    ce(this, "borderFills", /* @__PURE__ */ new Map());
  }
}
class Ql {
  constructor(r) {
    ce(this, "ctx");
    this.ctx = r;
  }
  parseHeader(r) {
    const n = new DOMParser().parseFromString(r, "text/xml");
    this.parseCharProperties(n), this.parseParaProperties(n), this.parseStyles(n), this.parseBorderFills(n);
  }
  parseCharProperties(r) {
    const e = r.getElementsByTagName("hh:charPr");
    for (let n = 0; n < e.length; n++) {
      const i = e[n], a = i.getAttribute("id") || "", s = parseInt(i.getAttribute("height") || "1000"), o = i.getElementsByTagNameNS(this.ctx.ns.hh, "bold")[0], h = o ? o.getAttribute("hangul") === "1" || o.getAttribute("latin") === "1" : !1, l = i.getElementsByTagNameNS(this.ctx.ns.hh, "italic")[0], f = l ? l.getAttribute("hangul") === "1" || l.getAttribute("latin") === "1" : !1, d = i.getElementsByTagNameNS(this.ctx.ns.hh, "underline")[0], c = ((d == null ? void 0 : d.getAttribute("type")) ?? "NONE") !== "NONE", g = i.getElementsByTagNameNS(this.ctx.ns.hh, "strikeout")[0], y = ((g == null ? void 0 : g.getAttribute("type")) ?? "NONE") !== "NONE", m = i.getElementsByTagNameNS(this.ctx.ns.hh, "superScript")[0], E = ((m == null ? void 0 : m.getAttribute("type")) ?? "NONE") !== "NONE", x = i.getElementsByTagNameNS(this.ctx.ns.hh, "subScript")[0], R = ((x == null ? void 0 : x.getAttribute("type")) ?? "NONE") !== "NONE";
      this.ctx.charProperties.set(a, {
        id: a,
        height: s,
        bold: h,
        italic: f,
        underline: c,
        strikeout: y,
        supscript: E,
        subscript: R
      });
    }
  }
  parseParaProperties(r) {
    const e = r.getElementsByTagName("hh:paraPr");
    for (let n = 0; n < e.length; n++) {
      const i = e[n], a = i.getAttribute("id") || "", s = i.getAttribute("align") || "LEFT";
      this.ctx.paraProperties.set(a, { id: a, align: s });
    }
  }
  parseStyles(r) {
    const e = r.getElementsByTagName("hh:style");
    for (let n = 0; n < e.length; n++) {
      const i = e[n], a = i.getAttribute("id") || "", s = i.getAttribute("type") || "PARA", o = parseInt(i.getAttribute("level") || "0"), h = i.getAttribute("paraPrIDRef") || "", l = i.getAttribute("charPrIDRef") || "";
      this.ctx.styles.set(a, { id: a, type: s, level: o, paraPrIDRef: h, charPrIDRef: l });
    }
  }
  parseBorderFills(r) {
    const e = Array.from(
      r.getElementsByTagNameNS(this.ctx.ns.hh, "borderFill")
    );
    for (const n of e) {
      const i = n.getAttribute("id");
      if (!i) continue;
      const a = (h) => {
        const l = n.getElementsByTagNameNS(this.ctx.ns.hh, h)[0];
        if (!l) return;
        const f = parseInt(l.getAttribute("width") || "0", 10), d = l.getAttribute("color") || "0", u = "#" + Number(d).toString(16).padStart(6, "0");
        return { width: f, color: u };
      }, s = {
        left: a("leftBorder"),
        right: a("rightBorder"),
        top: a("topBorder"),
        bottom: a("bottomBorder")
      }, o = n.getElementsByTagNameNS(this.ctx.ns.hh, "fillBrush")[0];
      if (o) {
        const h = o.getElementsByTagNameNS(this.ctx.ns.hh, "winBrush")[0];
        if (h) {
          const l = h.getAttribute("faceColor");
          l && (s.backgroundColor = "#" + Number(l).toString(16).padStart(6, "0"));
        }
      }
      this.ctx.borderFills.set(i, s);
    }
  }
}
class eh {
  constructor(r) {
    ce(this, "ctx");
    this.ctx = r;
  }
  convertSection(r) {
    const i = new DOMParser().parseFromString(r, "text/xml").documentElement;
    let a = "";
    const s = i.childNodes;
    for (let o = 0; o < s.length; o++) {
      const h = s[o];
      h.nodeType === 1 && h.localName === "p" && (a += this.convertParagraph(h));
    }
    return a.trim();
  }
  convertParagraph(r) {
    const e = r.getAttribute("styleIDRef") || "";
    let n = 0;
    if (e) {
      const l = this.ctx.styles.get(e);
      l && l.type === "HEADING" && l.level > 0 && (n = Math.min(l.level, 6));
    }
    let i = "", a = "";
    const o = Array.from(
      r.getElementsByTagNameNS(this.ctx.ns.hp, "tbl")
    ).filter((l) => {
      let f = l.parentNode;
      for (; f && f !== r; ) {
        if (f.localName === "tc") return !1;
        f = f.parentNode;
      }
      return !0;
    });
    for (const l of o)
      a += this.convertTable(l);
    const h = r.getElementsByTagNameNS(this.ctx.ns.hp, "run");
    for (const l of Array.from(h)) {
      if (l.parentNode !== r || l.getElementsByTagNameNS(this.ctx.ns.hp, "tbl").length > 0)
        continue;
      const f = l.getElementsByTagNameNS(this.ctx.ns.hp, "pic");
      if (f.length > 0) {
        for (const m of Array.from(f))
          i += this.convertImage(m);
        continue;
      }
      const d = l.getElementsByTagNameNS(this.ctx.ns.hp, "rect");
      if (d.length > 0) {
        for (const m of Array.from(d))
          i += this.convertRect(m);
        continue;
      }
      if (l.getElementsByTagNameNS(this.ctx.ns.hp, "secPr").length > 0) continue;
      const c = l.getAttribute("charPrIDRef") || "0", g = this.ctx.charProperties.get(c);
      let p = "";
      const y = l.getElementsByTagNameNS(this.ctx.ns.hp, "t");
      for (const m of Array.from(y))
        p += this.getTextContent(m);
      p && g && (p = this.applyInlineFormatting(p, g)), i += p;
    }
    return a ? a + (i.trim() ? i.trim() + `

` : "") : i.trim() ? n > 0 ? "#".repeat(n) + " " + i.trim() + `

` : i.trim() + `

` : `
`;
  }
  applyInlineFormatting(r, e) {
    var o, h;
    const n = ((o = r.match(/^(\s*)/)) == null ? void 0 : o[1]) ?? "", i = ((h = r.match(/(\s*)$/)) == null ? void 0 : h[1]) ?? "", a = r.trim();
    if (!a) return r;
    let s = a;
    return e.supscript ? `${n}<sup>${s}</sup>${i}` : e.subscript ? `${n}<sub>${s}</sub>${i}` : (e.strikeout && (s = `~~${s}~~`), e.underline && (s = `<u>${s}</u>`), e.bold && e.italic ? s = `***${s}***` : e.bold ? s = `**${s}**` : e.italic && (s = `*${s}*`), n + s + i);
  }
  getTextContent(r) {
    let e = "";
    for (const n of Array.from(r.childNodes))
      n.nodeType === 3 && (e += n.nodeValue);
    return e;
  }
  expandCells(r) {
    let e = Array.from(r.children).filter(
      (n) => n.localName === "tc"
    );
    return e.length || (e = Array.from(
      r.getElementsByTagNameNS(this.ctx.ns.hp, "tc")
    ).filter((n) => n.parentNode === r)), e;
  }
  convertTable(r) {
    let e = Array.from(r.children).filter(
      (o) => o.localName === "tr"
    );
    if (e.length || (e = Array.from(
      r.getElementsByTagNameNS(this.ctx.ns.hp, "tr")
    ).filter((o) => o.parentNode === r)), !e.length) return "";
    const n = this.expandCells(e[0]);
    let i = 0;
    for (const o of n) {
      const h = o.getElementsByTagNameNS(this.ctx.ns.hp, "cellSpan")[0];
      i += h ? parseInt(h.getAttribute("colSpan") || "1", 10) : 1;
    }
    const a = /* @__PURE__ */ new Set();
    let s = "<table><tbody>";
    for (let o = 0; o < e.length; o++) {
      const h = e[o], l = this.expandCells(h);
      s += "<tr>";
      let f = 0, d = 0;
      for (; f < i; ) {
        if (a.has(`${o},${f}`)) {
          f++;
          continue;
        }
        if (d < l.length) {
          const u = l[d++], c = u.getElementsByTagNameNS(this.ctx.ns.hp, "cellSpan")[0], g = c ? parseInt(c.getAttribute("colSpan") || "1", 10) : 1, p = c ? parseInt(c.getAttribute("rowSpan") || "1", 10) : 1, y = (g > 1 ? ` colspan="${g}"` : "") + (p > 1 ? ` rowspan="${p}"` : ""), m = this.extractCellText(u), A = this.getCellSizeStyle(u), E = this.getCellBorderStyle(u), x = [A, E].filter(Boolean).join(";");
          if (s += `<td${y}${x ? ` style="${x}"` : ""}>${m}</td>`, p > 1)
            for (let S = 1; S < p; S++)
              for (let R = 0; R < g; R++)
                a.add(`${o + S},${f + R}`);
          f += g;
        } else
          s += "<td></td>", f++;
      }
      s += "</tr>";
    }
    return s += "</tbody></table>", s;
  }
  getCellSizeStyle(r) {
    const e = r.getElementsByTagNameNS(this.ctx.ns.hp, "cellSz")[0];
    if (!e) return "";
    const n = e.getAttribute("width"), i = e.getAttribute("height"), a = [];
    if (n) {
      const s = Math.round(parseInt(n, 10) * 96 / 7200);
      a.push(`width:${s}px`);
    }
    if (i) {
      const s = Math.round(parseInt(i, 10) * 96 / 7200);
      a.push(`height:${s}px`);
    }
    return a.join(";");
  }
  getCellBorderStyle(r) {
    const e = r.getAttribute("borderFillIDRef");
    if (!e) return "";
    const n = this.ctx.borderFills.get(e);
    if (!n) return "";
    const i = [], a = (s, o) => {
      if (!o || o.width === 0) {
        i.push(`border-${s}:none`);
        return;
      }
      i.push(`border-${s}:${o.width}px solid ${o.color}`);
    };
    return a("left", n.left), a("right", n.right), a("top", n.top), a("bottom", n.bottom), n.backgroundColor && i.push(`background:${n.backgroundColor}`), i.join(";");
  }
  extractCellText(r) {
    const e = r.getElementsByTagNameNS(this.ctx.ns.hp, "tbl")[0];
    if (e)
      return this.convertTable(e);
    let n = r.getElementsByTagNameNS(this.ctx.ns.hp, "subList")[0] || r.getElementsByTagNameNS(this.ctx.ns.hp, "list")[0] || null;
    n || (n = r);
    const i = [], a = Array.from(
      n.getElementsByTagNameNS(this.ctx.ns.hp, "p")
    );
    for (const s of a) {
      let o = "";
      const h = Array.from(s.getElementsByTagNameNS(this.ctx.ns.hp, "run"));
      for (const l of h) {
        const f = this.ctx.charProperties.get(
          l.getAttribute("charPrIDRef") || "0"
        );
        let d = "";
        const u = l.getElementsByTagNameNS(this.ctx.ns.hp, "t");
        for (const c of Array.from(u))
          d += this.getTextContent(c);
        d && f && (d = this.applyInlineFormatting(d, f)), o += d;
      }
      o.trim() && i.push(o.trim());
    }
    return i.join("<br>");
  }
  //md에서는 미지원 예정
  convertImage(r) {
    const e = r.getElementsByTagNameNS(this.ctx.ns.hc, "img")[0];
    if (!e) return "";
    const n = e.getAttribute("binaryItemIDRef");
    if (!n) return "![image]()";
    const i = this.ctx.images.get(n);
    if (!i) return "![image]()";
    const a = i.ext === "png" ? "image/png" : i.ext === "jpg" || i.ext === "jpeg" ? "image/jpeg" : i.ext === "gif" ? "image/gif" : i.ext === "bmp" ? "image/bmp" : `image/${i.ext}`, s = this.uint8ArrayToBase64(i.data);
    return `![image](data:${a};base64,${s})`;
  }
  uint8ArrayToBase64(r) {
    let e = "";
    for (let i = 0; i < r.length; i += 32768)
      e += String.fromCharCode(...r.subarray(i, i + 32768));
    return btoa(e);
  }
  convertRect(r) {
    const e = r.getElementsByTagNameNS(
      this.ctx.ns.hp,
      "drawText"
    )[0];
    if (!e) return "";
    const n = e.getElementsByTagNameNS(
      this.ctx.ns.hp,
      "subList"
    )[0];
    if (!n) return "";
    let i = "";
    const a = Array.from(
      n.getElementsByTagNameNS(this.ctx.ns.hp, "p")
    ).filter((s) => s.parentNode === n);
    for (const s of a) {
      const o = Array.from(
        s.getElementsByTagNameNS(this.ctx.ns.hp, "run")
      ).filter((h) => h.parentNode === s);
      for (const h of o) {
        const l = h.getElementsByTagNameNS(this.ctx.ns.hp, "t");
        for (const f of Array.from(l))
          i += this.getTextContent(f);
      }
    }
    return i;
  }
}
class th {
  async convert(r) {
    const e = await yt.loadAsync(r);
    return await this.convertFromZip(e);
  }
  async convertFromZip(r) {
    const e = new ql(), n = new Ql(e), i = new eh(e), a = r.file("Contents/header.xml");
    a && n.parseHeader(await a.async("string"));
    const s = r.folder("BinData");
    if (s) {
      const l = [];
      s.forEach((f, d) => {
        d.dir || l.push({ name: f, file: d });
      });
      for (const { name: f, file: d } of l) {
        const u = await d.async("uint8array"), c = f.split("."), g = c.length > 1 ? c.pop().toLowerCase() : "", p = c.length > 0 ? c[0].split("/").pop() : f;
        e.images.set(p, { data: u, ext: g });
      }
    }
    const o = Object.keys(r.files).filter((l) => /Contents\/section\d+\.xml$/i.test(l)).sort((l, f) => {
      var c, g;
      const d = parseInt(((c = l.match(/section(\d+)/i)) == null ? void 0 : c[1]) ?? "0"), u = parseInt(((g = f.match(/section(\d+)/i)) == null ? void 0 : g[1]) ?? "0");
      return d - u;
    }), h = [];
    for (const l of o) {
      const f = r.file(l);
      if (f) {
        const d = await f.async("string"), u = i.convertSection(d);
        u && h.push(u);
      }
    }
    return h.join(`

`);
  }
}
async function rh(t) {
  try {
    return await new th().convert(t);
  } catch (r) {
    throw console.error("HWPX to MD conversion error:", r), new Error(`Failed to convert HWPX: ${r.message}`);
  }
}
const ph = () => {
  const [t, r] = Oe(!1), [e, n] = Oe(null), [i, a] = Oe(null);
  return { convert: async (o) => {
    r(!0), n(null), a(null);
    try {
      const h = await rh(o);
      return a(h), h;
    } catch (h) {
      throw n(h), h;
    } finally {
      r(!1);
    }
  }, isLoading: t, error: e, result: i };
};
class nh {
  async readFile(r) {
    return console.warn(`[BrowserFileSystem] readFile not supported in browser: ${r}`), new Uint8Array(0);
  }
  async exists(r) {
    return console.warn(`[BrowserFileSystem] exists check: ${r} -> false`), !1;
  }
  async writeFile(r, e) {
    console.log(`[BrowserFileSystem] Writing to ${r}, size: ${e.length}`);
  }
}
async function ah(t) {
  let r = "";
  if (typeof t == "string")
    r = t;
  else if (t instanceof File || t instanceof Blob)
    r = await t.text();
  else
    throw new Error("Invalid input type. Expected string, File, or Blob.");
  const e = new nh(), n = new Zi(e);
  try {
    const i = await n.convert(r);
    return new Blob([i], { type: "application/xml" });
  } catch (i) {
    throw console.error("MD to HWP conversion error", i), new Error(`Failed to convert MD to HWP: ${i.message}`);
  }
}
const uh = () => {
  const [t, r] = Oe(!1), [e, n] = Oe(null);
  return { convert: Si(async (a) => {
    r(!0), n(null);
    try {
      return await ah(a);
    } catch (s) {
      return console.error("MD to HWP Conversion failed:", s), n(s.message || "Conversion failed"), null;
    } finally {
      r(!1);
    }
  }, []), isConverting: t, error: e };
};
export {
  Zi as HmlConverter,
  cn as HwpToDocxConverter,
  ki as convertDocxToHwpx,
  Yl as convertHwpToDocx,
  Kl as convertHwpToMd,
  Di as convertHwpxToDocx,
  rh as convertHwpxToMd,
  ah as convertMdToHwp,
  lh as useDocxToHwpx,
  dh as useHwpToDocx,
  fh as useHwpToMd,
  hh as useHwpxToDocx,
  ph as useHwpxToMd,
  uh as useMdToHwp
};
