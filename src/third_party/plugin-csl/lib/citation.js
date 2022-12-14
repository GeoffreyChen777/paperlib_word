"use strict";

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = citation;

var _core = require("@citation-js/core");

var _engines = _interopRequireDefault(require("./engines.js"));

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

function citation(data, options = {}) {
  const {
    template = 'apa',
    lang,
    format = 'text'
  } = options;
  const ids = data.map(({
    id
  }) => id);
  const entries = options.entry ? [].concat(options.entry) : ids;
  data = _core.util.downgradeCsl(data);
  const citeproc = (0, _engines.default)(data, template, lang, format);
  citeproc.updateItems(ids);
  const {
    citationsPre = [],
    citationsPost = []
  } = options;

  const citation = citeproc.processCitationCluster({
    citationItems: entries.map(id => typeof id === 'object' ? id : {
      id,
    }),
    properties: {
      noteIndex: 0
    }
  }, citationsPre, citationsPost);
  return citation;
}