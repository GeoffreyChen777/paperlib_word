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
setCitePlugin()
const citations = new Cite()
let csl = "apa"

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
    if (platform === Office.PlatformType.Mac && version >= "15.32" || platform === Office.PlatformType.PC && version >= "16.12.7668.1000" || platform === Office.PlatformType.OfficeOnline) {

      webSocket = new WebSocket("wss://localhost.paperlib.app:21993");

      const setSocketEvent = () => {
        webSocket.onopen = (event) => {
          $('#connect-icon').css('display', 'none')
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
        Office.context.document.settings.set("csl", csl);
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

      Office.context.document.settings.set("citationsData", [])
      Office.context.document.settings.set("citationByIndex", [])
      csl = Office.context.document.settings.get("csl") || "apa"
      citations.add(Office.context.document.settings.get("citationsData") || [])
      const prepareEngine = require("@citation-js/plugin-csl/lib-mjs/engines").default;
      const cslEngine = prepareEngine(Office.context.document.settings.get("citationsData") || [], csl, 'en-US', 'text')
      const citationByIndex = Office.context.document.settings.get('citationByIndex') || []
      cslEngine.rebuildProcessorState(citationByIndex);

      $('#csl-style-select').val(csl)
    }

    if (platform === Office.PlatformType.Mac && version >= "16.64" || platform === Office.PlatformType.PC && version >= "22.08.15601.20148)" || platform === Office.PlatformType.OfficeOnline) {
      bookmarkAvailable = true;
    }
  }

});

async function handler(data: string) {
  const message = JSON.parse(data) as { type: string, response: any }
  switch (message.type) {
    case 'search':
      handleSearchResult(message.response)
      break;
  }
}


async function search(query: string) {
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

  Word.run(async (context) => {
    const range = context.document.getSelection();

    const contentControls = context.document.contentControls;
    contentControls.load('title,tag');
    await context.sync();

    const citationsPre = []
    const citationsPost = []

    for (const contentControl of contentControls.items) {
      const ccRange = contentControl.getRange('Whole');
      const compareResult = range.compareLocationWith(ccRange)
      await context.sync();

      if (!contentControl.title.startsWith('citation')) {
        continue
      }
      const citationID = contentControl.title.split(':')[1]

      if (compareResult.value === 'Before') {
        citationsPost.push([citationID, 0])
      } else {
        citationsPre.push([citationID, 0])
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
      const citationID = contentControl.title.split(':')[1]
      const newCitationString = citationIDStringMap[citationID]
      if (newCitationString) {
        contentControl.insertText(newCitationString, 'Replace')
      }
      citationIDStringMap[citationID] = null
    }

    const newCitationIDs = Object.keys(citationIDStringMap).filter((key) => {
      return citationIDStringMap[key] !== null
    })

    const content = range.insertText(citationIDStringMap[newCitationIDs[0]], 'Replace');
    if (bookmarkAvailable) {
      content.hyperlink = `#Bookmark_${ids[0]}`;
    }
    const contentControl = content.insertContentControl()
    contentControl.title = `citation:${newCitationIDs[0]}`;
    contentControl.tag = ids.join(';');

    await context.sync();

    const prepareEngine = require("@citation-js/plugin-csl/lib-mjs/engines").default;
    const cslEngine = prepareEngine([], csl, 'en-US', 'text')

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
          refRange.insertBookmark(reference[0]);
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