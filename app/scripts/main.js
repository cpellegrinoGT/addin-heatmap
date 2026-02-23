/**
 * Activity Heatmap — MyGeotab Add-In
 *
 * Renders driver activity (GPS trails or exception events) as a heatmap
 * using Leaflet + Leaflet.heat inside a MyGeotab custom page.
 */

geotab.addin.activityHeatmap = function () {
  "use strict";

  // ── Constants ──────────────────────────────────────────────────────────
  var MAX_POINTS = 100000;
  var MULTI_CALL_BATCH = 50;
  var GPS_MATCH_TOLERANCE_MS = 5 * 60 * 1000; // 5 minutes

  // ── State ──────────────────────────────────────────────────────────────
  var map, heatLayer, api;
  var cachedRules = [];
  var abortController = null;
  var firstFocus = true;

  // ── DOM refs (set during initialize) ───────────────────────────────────
  var els = {};

  // ── Helpers ────────────────────────────────────────────────────────────

  function $(id) {
    return document.getElementById(id);
  }

  function getDateRange() {
    var now = new Date();
    var preset = document.querySelector(".heatmap-preset.active");
    var key = preset ? preset.dataset.preset : "7days";
    var from, to;

    to = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);

    switch (key) {
      case "today":
        from = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        break;
      case "30days":
        from = new Date(now);
        from.setDate(from.getDate() - 30);
        from.setHours(0, 0, 0, 0);
        break;
      case "custom":
        from = els.fromDate.value ? new Date(els.fromDate.value + "T00:00:00") : new Date(now.getTime() - 7 * 86400000);
        to = els.toDate.value ? new Date(els.toDate.value + "T23:59:59") : to;
        break;
      case "7days":
      default:
        from = new Date(now);
        from.setDate(from.getDate() - 7);
        from.setHours(0, 0, 0, 0);
        break;
    }

    return { from: from.toISOString(), to: to.toISOString() };
  }

  function getSelectedDeviceIds() {
    var val = els.vehicle.value;
    if (val === "all") return null; // null means all
    return [{ id: val }];
  }

  function getSelectedRuleId() {
    return els.exceptionType.value;
  }

  function showLoading(show) {
    els.loading.style.display = show ? "flex" : "none";
    els.empty.style.display = "none";
  }

  function showEmpty(show) {
    els.empty.style.display = show ? "flex" : "none";
  }

  function setStats(count) {
    els.stats.textContent = count > 0 ? count.toLocaleString() + " points" : "";
  }

  function clearHeat() {
    if (heatLayer) {
      map.removeLayer(heatLayer);
      heatLayer = null;
    }
    setStats(0);
  }

  /** Evenly sample an array down to maxLen items. */
  function sampleArray(arr, maxLen) {
    if (arr.length <= maxLen) return arr;
    var stride = arr.length / maxLen;
    var out = [];
    for (var i = 0; i < maxLen; i++) {
      out.push(arr[Math.floor(i * stride)]);
    }
    return out;
  }

  /** Binary-search for the LogRecord nearest to targetTime. Returns null if outside tolerance. */
  function findNearestRecord(records, targetTime) {
    if (!records || records.length === 0) return null;
    var target = new Date(targetTime).getTime();
    var lo = 0, hi = records.length - 1;

    while (lo <= hi) {
      var mid = (lo + hi) >>> 1;
      var t = new Date(records[mid].dateTime).getTime();
      if (t < target) lo = mid + 1;
      else if (t > target) hi = mid - 1;
      else return records[mid]; // exact match
    }

    // lo is the insertion point — check neighbours
    var candidates = [];
    if (lo < records.length) candidates.push(records[lo]);
    if (lo - 1 >= 0) candidates.push(records[lo - 1]);

    var best = null, bestDiff = Infinity;
    candidates.forEach(function (r) {
      var diff = Math.abs(new Date(r.dateTime).getTime() - target);
      if (diff < bestDiff) { bestDiff = diff; best = r; }
    });

    return bestDiff <= GPS_MATCH_TOLERANCE_MS ? best : null;
  }

  /** Check if an AbortController has been aborted. */
  function isAborted() {
    return abortController && abortController.signal && abortController.signal.aborted;
  }

  // ── API Helpers ────────────────────────────────────────────────────────

  /** Wrap api.call in a Promise. */
  function apiCall(method, params) {
    return new Promise(function (resolve, reject) {
      api.call(method, params, resolve, reject);
    });
  }

  /** Wrap api.multiCall in a Promise. */
  function apiMultiCall(calls) {
    return new Promise(function (resolve, reject) {
      api.multiCall(calls, resolve, reject);
    });
  }

  /** Fetch devices, optionally filtered by group. */
  function loadDevices(groupFilter) {
    var search = {};
    if (groupFilter && groupFilter.length) {
      search.groups = groupFilter.map(function (g) { return { id: g.id || g }; });
    }
    return apiCall("Get", { typeName: "Device", search: search, resultsLimit: 5000 });
  }

  /** Fetch rules (for exception name matching). Cached after first call. */
  function loadRules() {
    if (cachedRules.length) return Promise.resolve(cachedRules);
    return apiCall("Get", { typeName: "Rule", resultsLimit: 5000 }).then(function (rules) {
      cachedRules = rules;
      return rules;
    });
  }

  /** Populate the vehicle dropdown. */
  function populateVehicles(devices) {
    var current = els.vehicle.value;
    els.vehicle.innerHTML = '<option value="all">All Vehicles</option>';
    devices.sort(function (a, b) { return (a.name || "").localeCompare(b.name || ""); });
    devices.forEach(function (d) {
      var opt = document.createElement("option");
      opt.value = d.id;
      opt.textContent = d.name || d.id;
      els.vehicle.appendChild(opt);
    });
    // restore selection if still valid
    if (current && els.vehicle.querySelector('option[value="' + current + '"]')) {
      els.vehicle.value = current;
    }
  }

  /** Populate the exception rule dropdown. */
  function populateRules(rules) {
    var current = els.exceptionType.value;
    els.exceptionType.innerHTML = '<option value="all">All Rules</option>';
    rules.sort(function (a, b) { return (a.name || "").localeCompare(b.name || ""); });
    rules.forEach(function (r) {
      var opt = document.createElement("option");
      opt.value = r.id;
      opt.textContent = r.name || r.id;
      els.exceptionType.appendChild(opt);
    });
    if (current && els.exceptionType.querySelector('option[value="' + current + '"]')) {
      els.exceptionType.value = current;
    }
  }

  // ── GPS Mode ───────────────────────────────────────────────────────────

  function fetchGpsData(deviceIds, dateRange) {
    // Build calls — one per device, or one for all
    var calls = [];

    if (!deviceIds) {
      // all vehicles — single call, rely on server limit
      calls.push(["Get", {
        typeName: "LogRecord",
        search: { fromDate: dateRange.from, toDate: dateRange.to },
        resultsLimit: MAX_POINTS
      }]);
    } else {
      deviceIds.forEach(function (d) {
        calls.push(["Get", {
          typeName: "LogRecord",
          search: { deviceSearch: { id: d.id }, fromDate: dateRange.from, toDate: dateRange.to },
          resultsLimit: MAX_POINTS
        }]);
      });
    }

    // Batch multiCalls
    var batches = [];
    for (var i = 0; i < calls.length; i += MULTI_CALL_BATCH) {
      batches.push(calls.slice(i, i + MULTI_CALL_BATCH));
    }

    return batches.reduce(function (chain, batch) {
      return chain.then(function (accumulated) {
        if (isAborted()) return accumulated;
        return apiMultiCall(batch).then(function (results) {
          results.forEach(function (r) {
            if (Array.isArray(r)) {
              accumulated = accumulated.concat(r);
            }
          });
          return accumulated;
        });
      });
    }, Promise.resolve([]));
  }

  function renderGpsHeatmap(records) {
    // Filter invalid GPS
    var points = records.filter(function (r) {
      return r.latitude !== 0 && r.longitude !== 0;
    });

    if (points.length === 0) {
      showEmpty(true);
      return;
    }

    points = sampleArray(points, MAX_POINTS);
    setStats(points.length);

    var latLngs = points.map(function (r) {
      return [r.latitude, r.longitude, 0.5]; // intensity
    });

    heatLayer = L.heatLayer(latLngs, {
      radius: 12,
      blur: 15,
      maxZoom: 17,
      gradient: { 0.2: "blue", 0.4: "cyan", 0.6: "lime", 0.8: "yellow", 1.0: "red" }
    }).addTo(map);

    // Fit bounds
    var bounds = L.latLngBounds(points.map(function (r) { return [r.latitude, r.longitude]; }));
    map.fitBounds(bounds, { padding: [30, 30] });
  }

  // ── Exception Mode ─────────────────────────────────────────────────────

  function fetchExceptionData(deviceIds, dateRange) {
    var ruleId = getSelectedRuleId();

    var exceptionSearch = {
      fromDate: dateRange.from,
      toDate: dateRange.to
    };
    if (deviceIds) {
      exceptionSearch.deviceSearch = { id: deviceIds[0].id };
    }
    if (ruleId !== "all") {
      exceptionSearch.ruleSearch = { id: ruleId };
    }

    // Fetch exceptions (server-side filtered by rule when a specific rule is selected)
    return apiCall("Get", {
      typeName: "ExceptionEvent",
      search: exceptionSearch,
      resultsLimit: 50000
    }).then(function (exceptions) {
      if (isAborted()) return [];
      if (exceptions.length === 0) return [];

      // Group exceptions by device
      var byDevice = {};
      exceptions.forEach(function (e) {
        var did = e.device ? e.device.id : null;
        if (!did) return;
        if (!byDevice[did]) byDevice[did] = [];
        byDevice[did].push(e);
      });

      // Fetch LogRecords per device for GPS matching
      var logCalls = Object.keys(byDevice).map(function (did) {
        return ["Get", {
          typeName: "LogRecord",
          search: { deviceSearch: { id: did }, fromDate: dateRange.from, toDate: dateRange.to },
          resultsLimit: 50000
        }];
      });

      var deviceIdList = Object.keys(byDevice);

      // Batch the log calls
      var batches = [];
      for (var i = 0; i < logCalls.length; i += MULTI_CALL_BATCH) {
        batches.push({ calls: logCalls.slice(i, i + MULTI_CALL_BATCH), ids: deviceIdList.slice(i, i + MULTI_CALL_BATCH) });
      }

      return batches.reduce(function (chain, batch) {
        return chain.then(function (accumulated) {
          if (isAborted()) return accumulated;
          return apiMultiCall(batch.calls).then(function (results) {
            results.forEach(function (records, idx) {
              var did = batch.ids[idx];
              if (!Array.isArray(records)) return;

              // Sort records by dateTime for binary search
              records.sort(function (a, b) {
                return new Date(a.dateTime).getTime() - new Date(b.dateTime).getTime();
              });

              // Match each exception to nearest GPS coord
              byDevice[did].forEach(function (exc) {
                var nearest = findNearestRecord(records, exc.activeFrom);
                if (nearest && nearest.latitude !== 0 && nearest.longitude !== 0) {
                  accumulated.push([nearest.latitude, nearest.longitude, 1.0]);
                }
              });
            });
            return accumulated;
          });
        });
      }, Promise.resolve([]));
    });
  }

  function renderExceptionHeatmap(points) {
    if (points.length === 0) {
      showEmpty(true);
      return;
    }

    points = sampleArray(points, MAX_POINTS);
    setStats(points.length);

    heatLayer = L.heatLayer(points, {
      radius: 18,
      blur: 22,
      maxZoom: 17,
      gradient: { 0.3: "yellow", 0.6: "orange", 1.0: "red" }
    }).addTo(map);

    var bounds = L.latLngBounds(points.map(function (p) { return [p[0], p[1]]; }));
    map.fitBounds(bounds, { padding: [30, 30] });
  }

  // ── Main Load ──────────────────────────────────────────────────────────

  function loadData() {
    // Cancel any in-flight request
    if (abortController) abortController.abort();
    abortController = new AbortController();

    clearHeat();
    showLoading(true);
    showEmpty(false);

    var dateRange = getDateRange();
    var deviceIds = getSelectedDeviceIds();
    var mode = els.datasource.value;

    var promise;
    if (mode === "gps") {
      promise = fetchGpsData(deviceIds, dateRange).then(function (records) {
        if (isAborted()) return;
        renderGpsHeatmap(records);
      });
    } else {
      promise = fetchExceptionData(deviceIds, dateRange).then(function (points) {
        if (isAborted()) return;
        renderExceptionHeatmap(points);
      });
    }

    promise.catch(function (err) {
      if (!isAborted()) {
        console.error("Activity Heatmap error:", err);
        showEmpty(true);
        els.empty.textContent = "Error loading data. Please try again.";
      }
    }).then(function () {
      if (!isAborted()) showLoading(false);
    });
  }

  // ── UI Event Handlers ──────────────────────────────────────────────────

  function onDatasourceChange() {
    var isException = els.datasource.value === "exceptions";
    els.exceptionTypes.style.display = isException ? "" : "none";
  }

  function onPresetClick(e) {
    var btn = e.target.closest(".heatmap-preset");
    if (!btn) return;

    document.querySelectorAll(".heatmap-preset").forEach(function (b) { b.classList.remove("active"); });
    btn.classList.add("active");

    var isCustom = btn.dataset.preset === "custom";
    els.customDates.style.display = isCustom ? "" : "none";

    if (isCustom && !els.fromDate.value) {
      // Default custom range to last 7 days
      var now = new Date();
      var from = new Date(now);
      from.setDate(from.getDate() - 7);
      els.fromDate.value = from.toISOString().slice(0, 10);
      els.toDate.value = now.toISOString().slice(0, 10);
    }
  }

  // ── Add-In Lifecycle ───────────────────────────────────────────────────

  return {
    initialize: function (freshApi, state, callback) {
      api = freshApi;

      // Cache DOM refs
      els.datasource = $("heatmap-datasource");
      els.exceptionTypes = $("heatmap-exception-types");
      els.exceptionType = $("heatmap-exception-type");
      els.vehicle = $("heatmap-vehicle");
      els.fromDate = $("heatmap-from");
      els.toDate = $("heatmap-to");
      els.customDates = $("heatmap-custom-dates");
      els.apply = $("heatmap-apply");
      els.stats = $("heatmap-stats");
      els.loading = $("heatmap-loading");
      els.empty = $("heatmap-empty");

      // Init Leaflet map
      map = L.map("heatmap-map").setView([39.8283, -98.5795], 4); // center of US
      L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
        attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
        maxZoom: 19
      }).addTo(map);

      // Event listeners
      els.datasource.addEventListener("change", onDatasourceChange);
      els.apply.addEventListener("click", loadData);
      document.querySelector(".heatmap-presets").addEventListener("click", onPresetClick);

      // Load vehicles + rules in parallel
      Promise.all([
        loadDevices(state.getGroupFilter()).then(populateVehicles),
        loadRules().then(populateRules)
      ]).then(function () {
        callback();
      }).catch(function (err) {
        console.error("Activity Heatmap init error:", err);
        callback();
      });
    },

    focus: function (freshApi, state) {
      api = freshApi;
      map.invalidateSize();

      // Refresh vehicle list with current group filter
      loadDevices(state.getGroupFilter()).then(populateVehicles).catch(function () {});

      // Auto-load on first focus
      if (firstFocus) {
        firstFocus = false;
        loadData();
      }
    },

    blur: function () {
      if (abortController) {
        abortController.abort();
        abortController = null;
      }
      showLoading(false);
    }
  };
};
