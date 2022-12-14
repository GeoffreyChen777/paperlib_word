import { util } from '@citation-js/core';
import prepareEngine from './engines.js';
export default function citation(data, options = {}) {
  const {
    template = 'apa',
    lang,
    format = 'text'
  } = options;
  const ids = data.map(({
    id
  }) => id);
  const entries = options.entry ? [].concat(options.entry) : ids;
  data = util.downgradeCsl(data);
  const citeproc = prepareEngine(data, template, lang, format);
  citeproc.updateItems(ids);
  const {
    citationsPre = [],
    citationsPost = []
  } = options;
  const citation = citeproc.processCitationCluster({
    citationItems: entries.map(id => typeof id === 'object' ? id : {
      id
    }),
    properties: {
      noteIndex: 0
    }
  }, citationsPre, citationsPost, format);
  return citation;
}