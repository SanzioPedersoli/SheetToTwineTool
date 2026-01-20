const TITLE = "";
const DRIVE_FOLDER_ID = "";
const AUTHOR = "";

const COL_QUEST = 'Quest';
const COL_ID = 'ID';
const COL_TRIGGER = 'Trigger';
const COL_CHARACTER = 'Character';
const COL_DIALOG = 'Dialog';
const COL_JUMPTO = 'JumpTo';
const COL_NOTES = 'Artist Notes';
const COL_MACROS = 'Designer Notes';

function exportDialoguesToTwee() {
  const values = getSheetValues();
  validateData(values);

  const headers = values[0];
  const rows = values.slice(1);
  const col = getHeaderMap(headers);
  const byQuest = groupByQuest(rows, col);

  const output = buildTweeOutput(byQuest, col);
  const path = saveTweeToDrive(output, TITLE+'.twee');
  return path;
}

function getSheetValues() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  return sheet.getDataRange().getValues();
}

function validateData(values) {
  if (!values || values.length < 2) {
    throw new Error('No data found.');
  }
}

function getHeaderMap(headers) {
  const col = {};
  headers.forEach(function (header, i) {
    col[header] = i;
  });
  return col;
}

function groupByQuest(rows, col) {
  const byQuest = {};
  const questIndex = col[COL_QUEST];

  rows.forEach(function (row) {
    const quest = row[questIndex] || 'NO_QUEST';
    if (!byQuest[quest]) byQuest[quest] = [];
    byQuest[quest].push(row);
  });

  return byQuest;
}

function buildTweeOutput(byQuest, col) {
  const passagesResult = buildPassages(byQuest, col);
  const linesWithStart = prependStartPassage(passagesResult.lines, passagesResult.firstPassageId);
  return linesWithStart.join('\n');
}

function buildPassages(byQuest, col) {
  const lines = [];
  let firstPassageId = null;

  const questIndex = col[COL_QUEST];
  const idIndex = col[COL_ID];
  const triggerIndex = col[COL_TRIGGER];
  const characterIndex = col[COL_CHARACTER];
  const dialogIndex = col[COL_DIALOG];
  const jumpToIndex = col[COL_JUMPTO];
  const notesIndex = col[COL_NOTES];
  const macrosIndex = col[COL_MACROS];

  const questNames = Object.keys(byQuest);
  const firstPassageByQuest = {};

  questNames.forEach(function (quest) {
    const data = sortRowsById(byQuest[quest], idIndex);
    byQuest[quest] = data;

    const firstRow = data.find(function (row) {
      const id = row[idIndex];
      return id !== '' && id !== null && id !== undefined;
    });

    if (firstRow) {
      firstPassageByQuest[quest] = getPassageId(quest, firstRow[idIndex]);
    }
  });

  questNames.forEach(function (quest) {
    const data = byQuest[quest];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const id = row[idIndex];

      if (id === '' || id === null || id === undefined) {
        continue;
      }

      const trigger = row[triggerIndex];
      const isChoiceRow = isChoice(trigger);
      const character = row[characterIndex];
      const dialog = row[dialogIndex];
      const jumpTo = row[jumpToIndex];
      const notes = notesIndex != null ? row[notesIndex] : '';
      const macros = macrosIndex != null ? row[macrosIndex] : '';

      const pid = getPassageId(quest, id);

      if (!firstPassageId) {
        firstPassageId = pid;
      }

      lines.push(':: ' + pid);

      let body = '';

      if (character) {
        if (character == 'DOCUMENT'){
          body += '<div class="document">' + dialog || '' + '</div>';
        } else {
          body += '<span class="character">' + character + ':</span> ';
          if (notes) {
            body += '<span class="tone">' + String(notes) + '</span> ';
          }
          body += dialog || '';
        }
      } else {
        let inner = '';
        if (notes) {
          inner += '<span class="tone">(' + String(notes) + ')</span> ';
        }
        inner += dialog || '';
        if (inner) {
          body += '<span class="narration">' + inner + '</span>';
        }
      }

      if (macros && String(macros).trim()) {
        if (body) {
          body += '\n';
        }
        body += String(macros);
      }

      lines.push(body);
      lines.push('');

      if (!isChoiceRow) {
        const choices = [];
        let j = i + 1;

        while (j < data.length && isChoice(data[j][triggerIndex])) {
          const choiceRow = data[j];
          const choiceTrigger = String(choiceRow[triggerIndex] || '');
          const label = choiceTrigger.replace(/\[.*?]/, '').trim();
          const choiceId = choiceRow[idIndex];
          const choiceMacros =
            macrosIndex != null ? choiceRow[macrosIndex] : '';

          if (choiceId !== '' && choiceId !== null && choiceId !== undefined) {
            const choicePassage = getPassageId(quest, choiceId);
            choices.push({
              label: label,
              target: choicePassage,
              macros: choiceMacros,
            });
          }

          j++;
        }

        const hasChoices = choices.length > 0;

        if (!hasChoices) {
          let target = null;

          if (jumpTo) {
            // Can jumop to quest ID number if in the same quest, full quest name if have to jump to another quest, like this: QQuestName_1234
            target = resolveJumpTarget(quest, jumpTo);
          } else {
            const nextRow = i + 1 < data.length ? data[i + 1] : null;

            if (nextRow) {
              const nextId = nextRow[idIndex];
              if (nextId !== '' && nextId !== null && nextId !== undefined) {
                target = resolveJumpTarget(quest, nextId);
              }
            } else {
              const questPos = questNames.indexOf(quest);
              const nextQuest = questNames[questPos + 1];
              if (nextQuest && firstPassageByQuest[nextQuest]) {
                target = firstPassageByQuest[nextQuest];
              }
            }
          }

          if (target) {
            lines.push('[[Next->' + target + ']]');
          }
        }

        choices.forEach(function (choice) {
          let line;
          const macroText =
            choice.macros != null ? String(choice.macros).trim() : '';

          if (macroText && macroText.charAt(0) === '(') {
            line =
              '* '+macroText +
              '[[[' +
              choice.label +
              '->' +
              choice.target +
              ']]]';
          } else {
            line = '* [[' + choice.label + '->' + choice.target + ']]';
          }

          lines.push(line);
        });
      } else {
        let target = null;

        if (jumpTo) {
          target = resolveJumpTarget(quest, jumpTo);
        } else {
          let k = i + 1;
          while (k < data.length) {
            const nextId = data[k][idIndex];
            if (nextId !== '' && nextId !== null && nextId !== undefined) {
              target = resolveJumpTarget(quest, nextId);
              break;
            }
            k++;
          }

          if (!target) {
            const questPos = questNames.indexOf(quest);
            const nextQuest = questNames[questPos + 1];
            if (nextQuest && firstPassageByQuest[nextQuest]) {
              target = firstPassageByQuest[nextQuest];
            }
          }
        }

        if (target) {
          lines.push('[[Next->' + target + ']]');
        }
      }

      lines.push('');
    }
  });

  return { lines: lines, firstPassageId: firstPassageId };
}

function prependStartPassage(lines, firstPassageId) {
  if (!firstPassageId) return lines;

  const header = [];

  header.push(':: StoryData');
  header.push('{');
  header.push(' "format": "Harlowe",');
  header.push(' "format-version": "3.3.9",');
  header.push(' "start": "Start",');
  header.push(' "zoom": 1');
  header.push('}');

  if (AUTHOR) {
    header.push(':: StoryAuthor');
    header.push(AUTHOR);
    header.push('');
  }

  if (TITLE) {
    header.push(':: StoryTitle');
    header.push(TITLE);
    header.push('');
  }

  header.push(':: Start');
  header.push('[[Begin->' + firstPassageId + ']]');
  header.push('');

  return header.concat(lines);
}

function saveTweeToDrive(output, fileName) {
  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const file = folder.createFile(fileName, output, MimeType.PLAIN_TEXT);
  const path = file.getUrl();
  Logger.log('Created Twee file: ' + path);
  return path;
}

function isChoice(trigger) {
  if (!trigger) return false;
  const value = String(trigger).toLowerCase();
  return value.indexOf('[choice]') !== -1;
}

function slugifyQuest(quest) {
  if (!quest) return 'Q';
  return (
    'Q' +
    String(quest)
      .replace(/\[[^\]]*]/g, '')
      .replace(/\s+/g, '')
      .replace(/[^A-Za-z0-9]/g, '')
  );
}

function getPassageId(quest, id) {
  return slugifyQuest(quest) + '_' + String(id).trim();
}

function resolveJumpTarget(currentQuest, jumpTo) {
  if (!jumpTo) return null;

  const raw = String(jumpTo).trim();

  if (/^Q[A-Za-z0-9]+_[0-9]+$/.test(raw)) {
    return raw;
  }

  const m = raw.match(/^(.+?)[#:](\d+)$/);
  if (m) {
    const otherQuest = m[1].trim();
    const otherId = m[2].trim();
    return getPassageId(otherQuest, otherId);
  }

  return getPassageId(currentQuest, raw);
}


function sortRowsById(rows, idIndex) {
  return rows.slice().sort(function (a, b) {
    const aID = Number(a[idIndex] || 0);
    const bID = Number(b[idIndex] || 0);
    return aID - bID;
  });
}
