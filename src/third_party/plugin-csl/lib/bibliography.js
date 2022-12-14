"use strict";

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = bibliography;

var _core = require("@citation-js/core");

var _engines = _interopRequireDefault(require("./engines.js"));

var _attr = require("./attr.js");

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

const getAffix = (source, affix) => typeof affix === 'function' ? affix(source) : affix || '';

function bibliography(data, options = {}) {
  const {
    template = 'apa',
    lang,
    format = 'text',
    nosort = false
  } = options;
  const ids = options.entry ? [].concat(options.entry) : data.map(({
    id
  }) => id);
  data = _core.util.downgradeCsl(data);
  const citeproc = (0, _engines.default)(data, template, lang, format);
  const sortedIds = citeproc.updateItems(ids, nosort);

  if (options.append || options.prepend) {
    const items = data.reduce((items, entry) => {
      items[entry.id] = entry;
      return items;
    }, {});

    citeproc.sys.wrapBibliographyEntry = function (id) {
      const entry = items[id];
      return [getAffix(entry, options.prepend), getAffix(entry, options.append)];
    };
  } else {
    citeproc.sys.wrapBibliographyEntry = () => ['', ''];
  }

  const bibliography = citeproc.makeBibliography();
  const [{
    bibstart,
    bibend
  }, bibBody] = bibliography;
  const entries = bibBody.map((element, index) => (0, _attr.getPrefixedEntry)(element, sortedIds[index]));

  if (options.asEntryArray) {
    return entries.map((element, index) => [sortedIds[index], element]);
  }

  return bibstart + entries.join('') + bibend;
}