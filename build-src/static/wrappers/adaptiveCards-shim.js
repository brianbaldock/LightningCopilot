/* global window */
(function() {
    // Preserve any existing AMD/CommonJS globals so we can restore later
    var originalDefine = window.define;
    var originalModule = window.module;
    var originalExports = window.exports;
    var originalGlobalThis = window.globalThis;
    var originalSelf = window.self;
    var originalWindow = window.window;

    try {
        // Ensure the UMD bundle sees a browser-like global object
        if (typeof window.globalThis === 'undefined') {
            window.globalThis = window;
        }
        if (typeof window.self === 'undefined') {
            window.self = window;
        }
        if (typeof window.window === 'undefined') {
            window.window = window;
        }
    } catch (e) {
        // Best-effort under Locker; ignore if assignments are blocked
    }

    // Temporarily disable AMD/CommonJS detection so the UMD assigns to window
    window.define = undefined;
    window.module = undefined;
    window.exports = undefined;

    try {
        (function() {
            /*__ADAPTIVE_CARDS_PAYLOAD__*/
        }).call(window);
    } finally {
        // Restore globals immediately after the payload executes
        window.define = originalDefine;
        window.module = originalModule;
        window.exports = originalExports;

        try {
            window.globalThis = originalGlobalThis || window;
            window.self = originalSelf || window;
            window.window = originalWindow || window;
        } catch (e) {
            // Ignore restoration failures under Locker constraints
        }
    }
})();
