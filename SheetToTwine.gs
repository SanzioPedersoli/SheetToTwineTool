function exportDialoguesToTwee() {
  const values = getSheetValues();
  validateData(values);

  const headers = values[0];
  const rows = values.slice(1);
  const colMap = getHeaderMap(headers);

  const storyModel = buildStoryModel(rows, colMap);
  const output = twineExporter(storyModel).join("\n");

  const path = saveTweeToDrive(output, TITLE + '.twee');
  return path;
}

function twineExporter(storyModel) {
  const lines = [];
  let firstPassageId = null;

  Object.keys(storyModel.quests).forEach(quest => {
    const passages = storyModel.quests[quest].passages;

    Object.keys(passages).forEach(id => {
      const passageId = getPassageId(quest, id);
      if (!firstPassageId) firstPassageId = passageId;

      const passage = passages[id];

      lines.push(':: ' + passageId);
      let body = buildPassageBody_(passage);
      lines.push(body);
      lines.push('');

      if (passage.next) pushNextLink_(lines, getPassageId(passage.next.quest, passage.next.id));
      pushChoiceLinks_(lines, passage.choices);
      lines.push('');
    });
  });

  return prependStartPassage(lines, firstPassageId);
}

function buildPassageBody_(passage) {
  let body = '';

  if (passage.character) {
    if (passage.character === 'DOCUMENT') {
      body += '<div class="document">' + (passage.dialog || '') + '</div>';
    } else {
      body += '<span class="character">' + passage.character + ':</span> ';
      if (passage.notes) body += '<span class="tone">' + passage.notes + '</span> ';
      body += passage.dialog || '';
    }
  } else {
    let inner = '';
    if (passage.notes) inner += '<span class="tone">(' + passage.notes + ')</span> ';
    inner += passage.dialog || '';
    if (inner) body += '<span class="narration">' + inner + '</span>';
  }

  if (passage.macros && String(passage.macros).trim()) {
    if (body) body += '\n';
    body += passage.macros;
  }

  return body;
}

function pushNextLink_(lines, targetId) {
  if (targetId) lines.push('[[Next->' + targetId + ']]');
}

function pushChoiceLinks_(lines, choices) {
  choices.forEach(c => {
    const macroText = c.macros ? String(c.macros).trim() : '';
    const line = macroText && macroText.charAt(0) === '('
      ? '* ' + macroText + '[[[' + c.label + '->' + getPassageId(c.target.quest, c.target.id) + ']]]'
      : '* [[' + c.label + '->' + getPassageId(c.target.quest, c.target.id) + ']]';
    lines.push(line);
  });
}

function getPassageId(quest, id) {
  return slugifyQuest(quest) + '_' + String(id).trim();
}

function slugifyQuest(quest) {
  if (!quest) return 'Q';
  return 'Q' + String(quest).replace(/\[[^\]]*]/g, '').replace(/\s+/g, '').replace(/[^A-Za-z0-9]/g, '');
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