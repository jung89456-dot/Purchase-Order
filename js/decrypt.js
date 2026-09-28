/*
 * 비밀번호가 걸린 엑셀(.xlsx, ECMA-376 Agile Encryption) 브라우저 내 복호화.
 * 스마트스토어 주문 엑셀처럼 암호가 걸린 파일을 서버 전송 없이 WebCrypto 로 푼다.
 *  - 컨테이너(CFB) 파싱: SheetJS 의 XLSX.CFB
 *  - 키 유도: H0 = H(salt + 비밀번호(UTF-16LE)), Hn = H(n + Hn-1) × spinCount, 블록키 결합
 *  - 본문: 4096바이트 세그먼트 단위 AES-CBC (IV = H(keySalt + 세그먼트번호))
 */
(function (root) {
  'use strict';

  var BLOCK_VERIFIER_INPUT = [0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79];
  var BLOCK_VERIFIER_VALUE = [0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e];
  var BLOCK_KEY = [0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6];
  var SEGMENT = 4096;

  function subtle() {
    var c = (typeof crypto !== 'undefined' && crypto.subtle) ? crypto : (typeof require === 'function' ? require('crypto').webcrypto : null);
    if (!c || !c.subtle) throw new Error('이 브라우저는 암호 해제를 지원하지 않습니다.');
    return c.subtle;
  }

  function concat() {
    var len = 0, i;
    for (i = 0; i < arguments.length; i++) len += arguments[i].length;
    var out = new Uint8Array(len), off = 0;
    for (i = 0; i < arguments.length; i++) { out.set(arguments[i], off); off += arguments[i].length; }
    return out;
  }

  function u32(n) { return new Uint8Array([n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255]); }

  function b64(s) {
    var bin = atob(s), out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  function utf16le(str) {
    var out = new Uint8Array(str.length * 2);
    for (var i = 0; i < str.length; i++) { var c = str.charCodeAt(i); out[i * 2] = c & 255; out[i * 2 + 1] = c >>> 8; }
    return out;
  }

  function fit(buf, len, padByte) {
    if (buf.length === len) return buf;
    if (buf.length > len) return buf.slice(0, len);
    var out = new Uint8Array(len); out.fill(padByte); out.set(buf); return out;
  }

  var HASH_NAMES = { SHA512: 'SHA-512', SHA384: 'SHA-384', SHA256: 'SHA-256', SHA1: 'SHA-1', 'SHA-1': 'SHA-1', 'SHA-256': 'SHA-256', 'SHA-384': 'SHA-384', 'SHA-512': 'SHA-512' };

  function hash(alg, data) {
    return subtle().digest(alg, data).then(function (b) { return new Uint8Array(b); });
  }

  function toBytes(content) {
    if (content instanceof Uint8Array) return content;
    if (Array.isArray(content)) return Uint8Array.from(content);
    return new Uint8Array(content);
  }

  function findEntry(cfb, name) {
    var e = root.XLSX ? root.XLSX.CFB.find(cfb, name) : null;
    return e ? toBytes(e.content) : null;
  }

  /**
   * 암호화된 OOXML(CFB 컨테이너)인지, 이 앱이 풀 수 있는 방식(Agile)인지 판별
   * @returns {{encrypted: boolean, supported: boolean}}
   */
  function inspect(data) {
    var sig = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
    for (var i = 0; i < 8; i++) if (data[i] !== sig[i]) return { encrypted: false, supported: false };
    try {
      var cfb = root.XLSX.CFB.read(data, { type: 'array' });
      var info = findEntry(cfb, 'EncryptionInfo');
      if (!info || !findEntry(cfb, 'EncryptedPackage')) return { encrypted: false, supported: false };
      var major = info[0] | (info[1] << 8), minor = info[2] | (info[3] << 8);
      return { encrypted: true, supported: major === 4 && minor === 4 };
    } catch (e) { return { encrypted: false, supported: false }; }
  }

  function isEncrypted(data) { return inspect(data).encrypted; }

  // 화면이 멈추지 않도록 반복 계산 사이사이 브라우저에 제어권을 넘긴다
  function yieldToUi(v) { return new Promise(function (res) { setTimeout(function () { res(v); }, 0); }); }

  function parseInfo(info) {
    var major = info[0] | (info[1] << 8), minor = info[2] | (info[3] << 8);
    if (!(major === 4 && minor === 4)) {
      var err = new Error('지원하지 않는 암호화 방식입니다(구형 엑셀 암호). 엑셀에서 비밀번호를 해제한 뒤 올려주세요.');
      err.code = 'UNSUPPORTED';
      throw err;
    }
    var xml = new TextDecoder('utf-8').decode(info.subarray(8));
    var doc = new DOMParser().parseFromString(xml, 'application/xml');
    var keyData = doc.getElementsByTagNameNS('*', 'keyData')[0];
    var encKey = doc.getElementsByTagNameNS('*', 'encryptedKey')[0];
    function a(el, n) { return el.getAttribute(n); }
    return {
      keyData: {
        salt: b64(a(keyData, 'saltValue')),
        hash: HASH_NAMES[a(keyData, 'hashAlgorithm')],
        blockSize: +a(keyData, 'blockSize'),
        keyBits: +a(keyData, 'keyBits')
      },
      key: {
        spin: +a(encKey, 'spinCount'),
        salt: b64(a(encKey, 'saltValue')),
        hash: HASH_NAMES[a(encKey, 'hashAlgorithm')],
        keyBits: +a(encKey, 'keyBits'),
        blockSize: +a(encKey, 'blockSize'),
        verifierInput: b64(a(encKey, 'encryptedVerifierHashInput')),
        verifierValue: b64(a(encKey, 'encryptedVerifierHashValue')),
        keyValue: b64(a(encKey, 'encryptedKeyValue'))
      }
    };
  }

  // WebCrypto AES-CBC 는 PKCS#7 패딩을 강제하므로, 패딩 블록을 만들어 붙인 뒤 복호화한다.
  function aesCbcNoPad(keyBytes, iv, data) {
    var s = subtle();
    return s.importKey('raw', keyBytes, { name: 'AES-CBC' }, false, ['encrypt', 'decrypt']).then(function (key) {
      var last = data.subarray(data.length - 16);
      var pad = new Uint8Array(16).fill(16);
      return s.encrypt({ name: 'AES-CBC', iv: last }, key, pad).then(function (enc) {
        var extra = new Uint8Array(enc).subarray(0, 16);
        return s.decrypt({ name: 'AES-CBC', iv: iv }, key, concat(data, extra));
      });
    }).then(function (b) { return new Uint8Array(b); });
  }

  function deriveBase(k, password) {
    return hash(k.hash, concat(k.salt, utf16le(password))).then(function (h) {
      // spinCount(보통 100,000) 반복 — 동기식 루프 대신 청크 단위 비동기 처리
      var i = 0;
      function step() {
        var chain = Promise.resolve(h);
        var end = Math.min(i + 1000, k.spin);
        for (; i < end; i++) {
          (function (n) { chain = chain.then(function (cur) { return hash(k.hash, concat(u32(n), cur)); }); })(i);
        }
        return chain.then(yieldToUi).then(function (cur) { h = cur; return i < k.spin ? step() : h; });
      }
      return step();
    });
  }

  function blockKey(k, base, block) {
    return hash(k.hash, concat(base, Uint8Array.from(block))).then(function (h) { return fit(h, k.keyBits / 8, 0x36); });
  }

  function arrEq(a, b) {
    if (a.length !== b.length) return false;
    for (var i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
    return true;
  }

  /**
   * 암호화된 xlsx 바이트 + 비밀번호 → 복호화된 xlsx 바이트
   * 비밀번호가 틀리면 code 'BAD_PASSWORD' 오류
   */
  function decrypt(data, password) {
    return Promise.resolve().then(function () {
      var cfb = root.XLSX.CFB.read(data, { type: 'array' });
      var info = parseInfo(findEntry(cfb, 'EncryptionInfo'));
      var pkg = findEntry(cfb, 'EncryptedPackage');
      var k = info.key, kd = info.keyData;
      return deriveBase(k, password).then(function (base) {
        return Promise.all([
          blockKey(k, base, BLOCK_VERIFIER_INPUT),
          blockKey(k, base, BLOCK_VERIFIER_VALUE),
          blockKey(k, base, BLOCK_KEY)
        ]);
      }).then(function (keys) {
        return Promise.all([
          aesCbcNoPad(keys[0], k.salt, k.verifierInput),
          aesCbcNoPad(keys[1], k.salt, k.verifierValue),
          aesCbcNoPad(keys[2], k.salt, k.keyValue)
        ]);
      }).then(function (res) {
        return hash(k.hash, res[0]).then(function (vh) {
          if (!arrEq(vh, res[1].subarray(0, vh.length))) {
            var err = new Error('비밀번호가 맞지 않습니다.');
            err.code = 'BAD_PASSWORD';
            throw err;
          }
          var secret = res[2].subarray(0, kd.keyBits / 8);
          var size = pkg[0] + pkg[1] * 256 + pkg[2] * 65536 + pkg[3] * 16777216 + (pkg[4] + pkg[5] * 256) * 4294967296;
          var body = pkg.subarray(8);
          var jobs = [];
          for (var seg = 0, off = 0; off < body.length; seg++, off += SEGMENT) {
            (function (seg, chunk) {
              jobs.push(hash(kd.hash, concat(kd.salt, u32(seg))).then(function (ivh) {
                return aesCbcNoPad(secret, fit(ivh, kd.blockSize, 0x36), chunk);
              }));
            })(seg, body.subarray(off, Math.min(off + SEGMENT, body.length)));
          }
          return Promise.all(jobs).then(function (parts) {
            return concat.apply(null, parts).subarray(0, size);
          });
        });
      });
    });
  }

  var api = { inspect: inspect, isEncrypted: isEncrypted, decrypt: decrypt };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.PODecrypt = api;
})(typeof self !== 'undefined' ? self : this);
