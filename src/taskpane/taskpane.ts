/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import jq from "jquery";
import { CSL } from "src/models/csl";
import { PaperEntity } from "../models/paper-entity";
import { formatString } from "../utils/string";

let bookmarkAvailable = false;
let webSocket: WebSocket;

var $ = jq

const Cite = require('citation-js')
const prepareEngine = require("@citation-js/plugin-csl/lib-mjs/engines").default;

setCitePlugin()
let citations = new Cite()
let csl = "apa"
let cslEngine;

let tempAddPaperEntities: Record<string, PaperEntity> = {}

function setCitePlugin() {
  const parseSingle = (paperEntityDraft: PaperEntity) => {
    let nameArray;
    if (paperEntityDraft.authors.includes(";")) {
      nameArray = paperEntityDraft.authors.split(";");
    } else {
      nameArray = paperEntityDraft.authors.split(",");
    }
    nameArray = nameArray.map((name) => {
      name = name.trim();
      const nameParts = name.split(" ");
      const given = nameParts.slice(0, nameParts.length - 1).join(" ");
      const family = nameParts[nameParts.length - 1];

      return {
        given: given,
        family: family,
      };
    });

    let citeKey = "";
    if (nameArray.length >= 1) {
      citeKey += nameArray[0].family.toLowerCase();
    }
    citeKey += paperEntityDraft.pubTime;
    const titleArray = paperEntityDraft.title.split(" ");
    for (const word of titleArray) {
      if (
        word.toLocaleLowerCase() !== "the" ||
        word.toLocaleLowerCase() !== "a"
      ) {
        citeKey += formatString({
          str: word.toLowerCase(),
          removeNewline: true,
          removeSymbol: true,
          removeWhite: true,
          trimWhite: true,
        });
        break;
      }
    }
    return {
      id: `${paperEntityDraft.id}`,
      type: ["article", "paper-conference", "article", "book"][
        paperEntityDraft.pubType
      ],
      "citation-key": citeKey,
      title: paperEntityDraft.title,
      author: nameArray,
      issued: {
        "date-parts": [[paperEntityDraft.pubTime]],
      },
      "container-title": paperEntityDraft.publication,
      publisher: paperEntityDraft.publisher,
      page: paperEntityDraft.pages,
      volume: paperEntityDraft.volume,
      issue: paperEntityDraft.number,
      DOI: paperEntityDraft.doi,
    };
  };

  const parseMulti = (paperEntityDrafts: PaperEntity[]) => {
    return paperEntityDrafts.map((paperEntityDraft) => {
      return parseSingle(paperEntityDraft);
    });
  };

  const predicateSingle = (paperEntityDraft: PaperEntity) => {
    return paperEntityDraft.codes !== undefined;
  };

  const predicateMulti = (paperEntityDrafts: PaperEntity[]) => {
    if (!!paperEntityDrafts?.[Symbol.iterator]) {
      return paperEntityDrafts.every((paperEntityDraft) => {
        return paperEntityDraft.codes !== undefined;
      });
    } else {
      return false;
    }
  };

  Cite.plugins.input.add("@paperlib/PaperEntity", {
    parse: parseSingle,
    parseType: {
      predicate: predicateSingle,
      dataType: "ComplexObject",
    },
  });
  Cite.plugins.input.add("@paperlib/PaperEntity[]", {
    parse: parseMulti,
    parseType: {
      predicate: predicateMulti,
      dataType: "Array",
    },
  });

  Cite.plugins.output.add("bibtex-key", (csls: CSL[]) => {
    return csls
      .map((csl) => {
        return csl["citation-key"];
      })
      .join(", ");
  });
}

Office.onReady((info) => {

  if (info.host === Office.HostType.Word) {

    const version = Office.context.diagnostics.version;
    const platform = Office.context.diagnostics.platform;
    console.log(version, platform);
    if (Office.context.requirements.isSetSupported('WordApi', "1.3")) {
      webSocket = new WebSocket("wss://localhost.paperlib.app:21993");

      const setSocketEvent = () => {
        webSocket.onopen = (event) => {
          $('#connect-icon').css('display', 'none')
          webSocket.send(JSON.stringify({ type: 'csl-names' }))
        };

        webSocket.onclose = (event) => {
          $('#connect-icon').css('display', 'block')
        };

        webSocket.onerror = (event) => {
          $('#connect-icon').css('display', 'block')
        };

        webSocket.onmessage = (event) => {
          handler(event.data)
        };

        if (webSocket.readyState === WebSocket.OPEN) {
          webSocket.send(JSON.stringify({ type: 'csl-names' }))
        }
      }

      setSocketEvent();

      $('#connect-icon').on('click', () => {
        console.log('reconnecting...')
        webSocket = new WebSocket("wss://localhost.paperlib.app:21993");
        setSocketEvent();
      })

      $('#search-bar').on('change', function (event) {
        search((event.target as HTMLInputElement).value)
      });

      $('#btn-add-ref').on('click', function () {
        insertOrRefreshReferences()
      });

      $('#csl-style-select').on('change', function (event) {
        csl = (event.target as HTMLSelectElement).value;
        if (['apa', 'vancouver', 'harvard1'].includes(csl)) {
          Office.context.document.settings.set("csl", csl);
          rebuildCitation();
        } else {
          webSocket.send(JSON.stringify({ type: 'load-csl', params: (event.target as HTMLSelectElement).value }))
        }
      });

      $('#btn-add-citation').on('click', function () {

        const existingIds = citations.data.map((item) => {
          return item.id;
        });

        for (const insertPaperEntity of Object.values(tempAddPaperEntities)) {
          const citation = new Cite(insertPaperEntity)

          if (existingIds.includes(insertPaperEntity.id)) {
            console.log("already exists");
          } else {
            citations.add(citation.data)
          }
        }
        Office.context.document.settings.set('citationsData', citations.data);
        Office.context.document.settings.saveAsync();

        insertCitation(new Cite(Object.values(tempAddPaperEntities)).getIds())
      });

    }

    if (Office.context.requirements.isSetSupported('WordApi', "1.4")) {
      bookmarkAvailable = true;
    }
  }

});

async function handler(data: string) {
  console.log('handler')
  const message = JSON.parse(data) as { type: string, response: any }
  switch (message.type) {
    case 'search':
      handleSearchResult(message.response)
      break;
    case 'csl-names':
      handleCSLNamesResult(message.response)
      break;
    case 'load-csl':
      handleCSLResult(message.response)
      break;
  }
}


async function search(query: string) {
  console.log('search')
  for (const paperEntity of Object.values(tempAddPaperEntities)) {
    $(`#check-${paperEntity.id}`).css('visibility', 'hidden')
    delete tempAddPaperEntities[`${paperEntity.id}`]
    if (Object.keys(tempAddPaperEntities).length === 0) {
      $('#btn-add-citation').css('display', 'none')
    }
  }
  webSocket.send(JSON.stringify({ type: 'search', params: { query: query } }))
}

let searchResults = {};
async function handleSearchResult(results: { id: string, title: string, authors: string, publication: string, pubTime: string }[]) {
  console.log('handle search result')
  $('#results-container').empty()
  for (const result of results) {
    searchResults[result.id] = result
  }
  results.forEach((item) => {
    $('#results-container').append(`
    <div class="result-item-container">
      <div class="result-check-container" id="check-${item.id}">
        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" class="bi bi-check-lg" viewBox="0 0 16 16">
          <path d="M12.736 3.97a.733.733 0 0 1 1.047 0c.286.289.29.756.01 1.05L7.88 12.01a.733.733 0 0 1-1.065.02L3.217 8.384a.757.757 0 0 1 0-1.06.733.733 0 0 1 1.047 0l3.052 3.093 5.4-6.425a.247.247 0 0 1 .02-.022Z"/>
        </svg>
      </div>
      <div class="result-container" id="${item.id}">
        <div class="result-title">${item.title}</div>
        <div class="result-authors">${item.authors}</div>
        <div class="result-pub">
            <div class="result-pub-year">${item.pubTime}</div>
            <div class="result-pub-divider">|</div>
            <div class="result-pub-title">${item.publication}</div>
        </div>
      </div>
    </div>
    <hr class="result-divider" />
    `)
  })

  $('.result-container').on('click', function (event) {
    const id = (event.target as HTMLElement).id;
    if (id) {
      const insertPaperEntity = searchResults[id];

      if (tempAddPaperEntities[id]) {
        delete tempAddPaperEntities[id]
        $(`#check-${id}`).css('visibility', 'hidden')
        if (Object.keys(tempAddPaperEntities).length === 0) {
          $('#btn-add-citation').css('display', 'none')
        }
      } else {
        tempAddPaperEntities[id] = insertPaperEntity;
        $(`#check-${id}`).css('visibility', 'visible')
        $("#btn-add-citation").css('display', 'flex')
      }
    }
  })
}

function insertCitation(ids: string[]) {
  console.log('insert citation')
  Word.run(async (context) => {
    const range = context.document.getSelection();
    const ccRange = range.insertContentControl();
    ccRange.tag = ids.join(';');
    ccRange.title = 'citation:new';

    const contentControls = context.document.contentControls;
    contentControls.load('title,tag');
    await context.sync();

    const citationsPre = []
    const citationsPost = []

    let insertPost = false
    for (const contentControl of contentControls.items) {
      if (!contentControl.title.startsWith('citation')) {
        continue
      }
      const citationID = contentControl.title.split(':')[1]
      if (citationID === 'new') {
        insertPost = true
      } else {
        if (insertPost) {
          citationsPost.push([citationID, 0])
        } else {
          citationsPre.push([citationID, 0])
        }
      }
    }

    const citationObjs = citations.format('citation', {
      format: 'text',
      template: csl,
      lang: 'en-US',
      entry: ids,
      citationsPre,
      citationsPost,
    })[1] as [number, string, string][];

    const citationIDStringMap = {}
    for (const citationObj of citationObjs) {
      citationIDStringMap[citationObj[2]] = citationObj[1]
    }

    for (const contentControl of contentControls.items) {
      if (!contentControl.title.startsWith('citation')) {
        continue
      }
      const citationID = contentControl.title.split(':')[1]
      if (citationID !== 'new') {
        const paperIDs = contentControl.tag.split(';')
        const newCitationString = citationIDStringMap[citationID]
        if (newCitationString) {
          const content = contentControl.insertText(newCitationString, 'Replace')
          if (bookmarkAvailable) {
            content.hyperlink = `#Bookmark_${paperIDs[0]}`;
          }
        }
        delete citationIDStringMap[citationID]
      }
    }

    const newCitationIDs = Object.keys(citationIDStringMap).filter((key) => {
      return citationIDStringMap[key] !== null
    })

    ccRange.title = `citation:${newCitationIDs[0]}`;
    const content = ccRange.insertText(citationIDStringMap[newCitationIDs[0]], 'Replace');
    if (bookmarkAvailable) {
      content.hyperlink = `#Bookmark_${ids[0]}`;
    }

    await context.sync();

    Office.context.document.settings.set('citationByIndex', cslEngine.registry.citationreg.citationByIndex);
    Office.context.document.settings.saveAsync();

    for (const paperEntity of Object.values(tempAddPaperEntities)) {
      $(`#check-${paperEntity.id}`).css('visibility', 'hidden')
      delete tempAddPaperEntities[`${paperEntity.id}`]
      if (Object.keys(tempAddPaperEntities).length === 0) {
        $('#btn-add-citation').css('display', 'none')
      }
    }
  });

}

function insertOrRefreshReferences() {
  console.log('insert or refresh references')
  Word.run(async (context) => {
    const contentControls = context.document.contentControls;
    contentControls.load('title,tag');
    await context.sync();

    const tags = contentControls.items.filter(item => item.title.startsWith('citation') && item.tag).map((item) => item.tag?.split(';')).flat();
    const uniqueIds = new Set(tags);
    const ids = Array.from(uniqueIds);

    const referenceString = citations.format('bibliography', {
      format: 'text',
      template: csl,
      lang: 'en-US',
      entry: ids,
      prepend(entry: PaperEntity) { return `${entry.id}:` }
    });

    const references = []

    for (const refStr of referenceString.split('\n')) {
      const id = refStr.split(':')[0]
      const ref = refStr.split(':').slice(1).join(':')
      references.push([id, ref])
    }

    const existingReferences = context.document.body.contentControls.getByTitle('references');
    existingReferences.load('title,tag');
    await context.sync();

    if (existingReferences.items.length > 0) {
      existingReferences.items[0].clear();

      for (const reference of references) {
        const refRange = existingReferences.items[0].insertText(`${reference[1]}\n`, Word.InsertLocation.end);
        if (bookmarkAvailable) {
          refRange.insertBookmark(`Bookmark_${reference[0]}`);
        }
      }
    } else {
      const range = context.document.getSelection();
      const contentControl = range.insertContentControl()
      contentControl.title = 'references';

      for (const reference of references) {
        const refRange = contentControl.insertText(`${reference[1]}
`, Word.InsertLocation.end);

        if (bookmarkAvailable) {
          refRange.insertBookmark(`Bookmark_${reference[0]}`);
        }
        await context.sync();
      }
    }
  })
}

function handleCSLNamesResult(results: Array<string>) {
  console.log('handle csl names result')
  const currentCSL = Office.context.document.settings.get('csl') || 'apa';
  for (const key of results) {
    $('#csl-style-select').append(`<option value="${key}">${key}</option>`)
  }

  if (results.includes(currentCSL) || ['apa', 'vancouver', 'harvard1'].includes(currentCSL)) {
    $('#csl-style-select').val(currentCSL);
    csl = currentCSL;
  } else {
    csl = 'apa';
    Office.context.document.settings.set('csl', 'apa');
  }

  console.log("Load", csl)

  if (!['apa', 'vancouver', 'harvard1'].includes(csl)) {
    webSocket.send(JSON.stringify({
      type: 'load-csl',
      params: Office.context.document.settings.get('csl')
    }))
  } else {
    rebuildCitation();
  }
}

function handleCSLResult(result: string) {
  console.log('handle csl result')
  if (result) {
    csl = $('#csl-style-select').val() as string;
    Office.context.document.settings.set('csl', $(`#csl-style-select`).val());
    Office.context.document.settings.saveAsync();

    let config = Cite.plugins.config.get("@csl");
    config.templates.add(csl, result);
    rebuildCitation();
  } else {
    $('#csl-style-select').val(Office.context.document.settings.get('csl'));
  }
}


function rebuildCitation() {
  console.log('rebuild citation')
  const citationsData = Office.context.document.settings.get("citationsData") || [];
  citations = null;
  citations = new Cite(citationsData);
  cslEngine = null;
  cslEngine = prepareEngine(citationsData, csl, 'en-US', 'text');
  const citationByIndex = Office.context.document.settings.get('citationByIndex') || [];
  cslEngine.rebuildProcessorState(citationByIndex);
}