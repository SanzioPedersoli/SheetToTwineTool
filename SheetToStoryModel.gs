function exportDialoguesToGenericNarrativeModel() {
  const values = getSheetValues();
  validateData(values);

  const headers = values[0];
  const rows = values.slice(1);
  const colMap = getHeaderMap(headers);

  const output = buildStoryModel(rows, colMap);

  const path = saveTweeToDrive(JSON.stringify(output, null, "  "), TITLE + '.json');
  return path;
}

function buildStoryModel(rows, colMap) {
  const ctx = buildColumnContext_(colMap);
  const byQuest = {};

  rows.forEach(row => {
    const questName = readCell_(row, ctx.questIndex, 'NO_QUEST');
    if (!byQuest[questName]) byQuest[questName] = [];
    byQuest[questName].push(row);
  });

  const storyModel = { quests: {} };

  Object.keys(byQuest).forEach(quest => {
    const sortedRows = sortRowsById(byQuest[quest], ctx.idIndex);
    storyModel.quests[quest] = { passages: {} };

    sortedRows.forEach((row, i) => {
      const id = readCell_(row, ctx.idIndex);
      const nextRow = sortedRows[i+1];
      if (isBlank_(id)) return;

      const passage = buildPassageModel_(row, nextRow, ctx, sortedRows, i);
      storyModel.quests[quest].passages[id] = passage;
    });
  });

  return storyModel;
}

function buildColumnContext_(colMap) {
  return {
    questIndex: colMap[COL_QUEST],
    idIndex: colMap[COL_ID],
    triggerIndex: colMap[COL_TRIGGER],
    characterIndex: colMap[COL_CHARACTER],
    dialogIndex: colMap[COL_DIALOG],
    jumpToIndex: colMap[COL_JUMPTO],
    notesIndex: colMap[COL_NOTES],
    macrosIndex: colMap[COL_MACROS],
  };
}

function buildPassageModel_(row, nextRow, ctx, data, idx) {
  const character = readCell_(row, ctx.characterIndex);
  const dialog = readCell_(row, ctx.dialogIndex);
  const notes = readCell_(row, ctx.notesIndex);
  const macros = readCell_(row, ctx.macrosIndex);
  const trigger = readCell_(row, ctx.triggerIndex);
  const jumpTo = readCell_(row, ctx.jumpToIndex);

  const isChoiceRow = isChoice(trigger);
  const choices = isChoiceRow ? [] : collectFollowingChoices_(data, idx, ctx);

  let next = null;
  next = resolveJumpTargetAgnostic_(row, nextRow, ctx, jumpTo);

  return {
    character,
    dialog,
    notes,
    macros,
    choices,
    next,
    trigger,
  };
}

function collectFollowingChoices_(data, i, ctx) {
  const choices = [];
  let j = i + 1;

  while (j < data.length && isChoice(data[j][ctx.triggerIndex])) {
    const row = data[j];
    const label = String(row[ctx.triggerIndex] || '').replace(/\[.*?\]/, '').trim();
    const choiceId = readCell_(row, ctx.idIndex);
    const macros = readCell_(row, ctx.macrosIndex);

    if (!isBlank_(choiceId)) {
      choices.push({
        label,
        target: { quest: readCell_(row, ctx.questIndex), id: choiceId },
        macros,
      });
    }
    j++;
  }

  return choices;
}

function resolveJumpTargetAgnostic_(row, nextRow, ctx, jumpTo) {
  const currentQuest = readCell_(row, ctx.questIndex);

  if (nextRow && !isChoice(row[ctx.triggerIndex]) && isChoice(nextRow[ctx.triggerIndex])) {
    return null;
  }

  if (!jumpTo) {
    if (!nextRow) return null;
    const nextRowQuest = readCell_(nextRow, ctx.questIndex);
    const nextRowId = readCell_(nextRow, ctx.idIndex);    
    return { quest: nextRowQuest, id: nextRowId };
  }
  else{
    const raw = String(jumpTo).trim();

    const match = raw.match(/(?<quest>[\d\w]*)(?:_)(?<id>\d.*)/);
    if (match && match.length==3)
      return { quest: match[1].trim(), id: match[2].trim() };
    else
      return { quest: currentQuest, id: raw };
  }
}

function sortRowsById(rows, idIndex) {
  return rows.slice().sort((a, b) => (Number(a[idIndex] || 0) - Number(b[idIndex] || 0)));
}

function isChoice(trigger) {
  if (!trigger) return false;
  return String(trigger).toLowerCase().includes('[choice]');
}
