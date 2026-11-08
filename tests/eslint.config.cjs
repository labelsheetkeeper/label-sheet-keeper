module.exports = [{
  files: ["**/*.js"],
  languageOptions: {
    ecmaVersion: 2018, sourceType: "script",
    globals: { window:"readonly", document:"readonly", Office:"readonly", Word:"readonly", localStorage:"readonly", navigator:"readonly", console:"readonly",
      location:"readonly", URL:"readonly", FileReader:"readonly", Image:"readonly", DOMParser:"readonly", XMLSerializer:"readonly", fetch:"readonly", AbortController:"readonly", File:"readonly", Event:"readonly", setTimeout:"readonly", clearTimeout:"readonly",
      setInterval:"readonly", clearInterval:"readonly", Promise:"readonly", Uint8Array:"readonly", globalThis:"readonly", module:"writable", JSON:"readonly", Math:"readonly", Date:"readonly", Number:"readonly", String:"readonly", Object:"readonly", Array:"readonly", isFinite:"readonly", Infinity:"readonly", decodeURIComponent:"readonly", Error:"readonly", RegExp:"readonly" }
  },
  rules: { "no-undef": "error", "no-unused-vars": ["warn", { "args": "none" }], "no-redeclare": "error", "no-dupe-keys": "error", "no-unreachable": "error", "no-use-before-define": ["error", { "functions": false, "variables": true }] }
}];
