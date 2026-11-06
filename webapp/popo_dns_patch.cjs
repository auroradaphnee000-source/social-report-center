"use strict";

// The service platform normally injects an /etc/hosts entry for this one
// internal tunnel. Some container recreations have received the gateway URL
// without the host entry. Patch DNS only inside the Fabric CLI subprocess so
// TLS still sees and validates the original hostname.
const dns = require("node:dns");

const targetHost = "cli.popo.netease.com";
const tunnelIp = String(process.env.POPO_TUNNEL_IP || "").trim();

if (tunnelIp) {
  const originalLookup = dns.lookup.bind(dns);

  dns.lookup = function patchedLookup(hostname, options, callback) {
    let lookupOptions = options;
    let done = callback;

    if (typeof lookupOptions === "function") {
      done = lookupOptions;
      lookupOptions = undefined;
    }
    if (hostname !== targetHost || typeof done !== "function") {
      return originalLookup(hostname, options, callback);
    }

    const normalized =
      typeof lookupOptions === "object" && lookupOptions !== null
        ? lookupOptions
        : {};
    if (normalized.all) {
      process.nextTick(done, null, [{ address: tunnelIp, family: 4 }]);
    } else {
      process.nextTick(done, null, tunnelIp, 4);
    }
  };

  const promiseApi = dns.promises;
  if (promiseApi && typeof promiseApi.lookup === "function") {
    const originalPromiseLookup = promiseApi.lookup.bind(promiseApi);
    promiseApi.lookup = async function patchedPromiseLookup(hostname, options) {
      if (hostname !== targetHost) {
        return originalPromiseLookup(hostname, options);
      }
      const normalized =
        typeof options === "object" && options !== null ? options : {};
      return normalized.all
        ? [{ address: tunnelIp, family: 4 }]
        : { address: tunnelIp, family: 4 };
    };
  }
}
