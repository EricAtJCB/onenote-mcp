'use strict';
// Constant-time Buffer comparison. Drop-in for buffer-equal-constant-time,
// which reads SlowBuffer at load time and crashes on Node 25 and newer.
var crypto = require('crypto');

function bufferEq(a, b) {
  if (!Buffer.isBuffer(a) || !Buffer.isBuffer(b)) return undefined;
  if (a.length !== b.length) return false;
  return crypto.timingSafeEqual(a, b);
}

// The original package exposes these; callers here never use them.
bufferEq.install = function () {};
bufferEq.restore = function () {};

module.exports = bufferEq;
