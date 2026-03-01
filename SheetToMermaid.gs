function exportDialoguesToMermaid() {
  const values = getSheetValues();
  validateData(values);

  const headers = values[0];
  const rows = values.slice(1);
  const colMap = getHeaderMap(headers);

  const data = buildStoryModel(rows, colMap);

  const output = narrativeToMermaid_(data, {
    direction: 'TD',
    maxDialogChars: 200,
    includeMissingTargets: true
  });

  const path = saveToDrive(output, TITLE, ".txt");
  return path;
}

function narrativeToMermaid_(data, options) {
  options = Object.assign({
    direction: 'TD',
    maxDialogChars: 200,
    includeMissingTargets: true
  }, options || {});

  if (!data || !data.quests) throw new Error('Invalid narrative JSON: missing "quests".');

  const questInfoByName = new Map();
  const questLinesBySlug = new Map();
  const edgeLines = [];
  const classLines = [];
  const definedNodes = new Set();

  const entryNodeByQuestNameAndId = new Map();
  const exitNodeByQuestNameAndId = new Map();
  const missingTargetsByQuestNameAndId = new Map();

  const questNames = prepareQuestMetadata_(data, questInfoByName, questLinesBySlug);

  createPassages_(questNames, data, questInfoByName, questLinesBySlug, options, definedNodes, classLines, edgeLines, entryNodeByQuestNameAndId, exitNodeByQuestNameAndId);
  manageNavigation_(questNames, data, questInfoByName, questLinesBySlug, exitNodeByQuestNameAndId, entryNodeByQuestNameAndId, missingTargetsByQuestNameAndId, definedNodes, options, edgeLines);

  const lines = [];
  lines.push(`flowchart ${options.direction}`);
  lines.push(`classDef narrator fill:#FFEB3B,stroke:#333,color:#000;`);


  createQuestsGroups_(questNames, questInfoByName, questLinesBySlug, lines, options, missingTargetsByQuestNameAndId);

  lines.push('');
  edgeLines.forEach((e) => lines.push(e));
  lines.push('');
  classLines.forEach((c) => lines.push(c));

  return lines.join('\n');
}

function createQuestsGroups_(questNames, questInfoByName, questLinesBySlug, lines, options, missingTargetsByQuestNameAndId) {
  questNames.forEach((qName) => {
    const qi = questInfoByName.get(qName);
    const questLines = questLinesBySlug.get(qi.slug);

    lines.push(`subgraph ${qi.subgraphId}["${escapeNodeLabel_(qi.title)}"]`);
    lines.push(` direction ${options.direction}`);
    questLines.forEach((l) => lines.push(` ${l}`));
    lines.push(`end`);
  });

  // Missing targets group
  if (options.includeMissingTargets && missingTargetsByQuestNameAndId.size > 0) {
    lines.push(`subgraph __missing["(missing targets)"]`);
    lines.push(` direction ${options.direction}`);
    const missingBlock = questLinesBySlug.get('__missing');
    if (missingBlock) missingBlock.forEach((l) => lines.push(` ${l}`));
    lines.push(`end`);
  }
}

function createPassages_(questNames, data, questInfoByName, questLinesBySlug, options, definedNodes, classLines, edgeLines, entryNodeByQuestNameAndId, exitNodeByQuestNameAndId) {
  questNames.forEach((qName) => {
    const quest = data.quests[qName];
    const qi = questInfoByName.get(qName);
    const questLines = questLinesBySlug.get(qi.slug);

    if (!quest || !quest.passages) return;

    Object.keys(quest.passages).forEach((pidRaw) => {
      const pid = String(pidRaw);
      const passage = quest.passages[pidRaw];

      const baseId = makePassageNodeId_(pid);
      const passageKey = makePassageKey_(qName, pid);

      createPassageNode_(passage, pid, options, questLines, definedNodes, baseId, classLines);
      manageMacroNodes_(passage, baseId, questLines, definedNodes, edgeLines, entryNodeByQuestNameAndId, passageKey, exitNodeByQuestNameAndId);
    });
  });
}

function manageNavigation_(questNames, data, questInfoByName, questLinesBySlug, exitNodeByQuestNameAndId, entryNodeByQuestNameAndId, missingTargetsByQuestNameAndId, definedNodes, options, edgeLines) {
  questNames.forEach((qName) => {
    const quest = data.quests[qName];
    const qi = questInfoByName.get(qName);
    const questLines = questLinesBySlug.get(qi.slug);

    if (!quest || !quest.passages) return;

    Object.keys(quest.passages).forEach((pidRaw) => {
      const pid = String(pidRaw);
      const passage = quest.passages[pidRaw];
      const passageKey = makePassageKey_(qName, pid);
      const fromExit = exitNodeByQuestNameAndId.get(passageKey);

      manageChoices_(passage, data, entryNodeByQuestNameAndId, missingTargetsByQuestNameAndId, questLinesBySlug, definedNodes, options, qi, pid, questLines, edgeLines, fromExit);
      manageNextButton_(passage, data, entryNodeByQuestNameAndId, missingTargetsByQuestNameAndId, questLinesBySlug, definedNodes, options, edgeLines, fromExit);
    });
  });
}

function manageNextButton_(passage, data, entryNodeByQuestNameAndId, missingTargetsByQuestNameAndId, questLinesBySlug, definedNodes, options, edgeLines, fromExit) {
  if (passage.next && passage.next.quest != null && passage.next.id != null) {
    const tgtQuest = passage.next.quest;
    const tgtId = String(passage.next.id);

    const toEntry = resolveTargetEntry_(
      tgtQuest, tgtId,
      data, entryNodeByQuestNameAndId,
      missingTargetsByQuestNameAndId, questLinesBySlug, definedNodes, options
    );

    edgeLines.push(`${fromExit} --> ${toEntry}`);
  }
}

function manageChoices_(passage, data, entryNodeByQuestNameAndId, missingTargetsByQuestNameAndId, questLinesBySlug, definedNodes, options, qi, pid, questLines, edgeLines, fromExit) {
  if (Array.isArray(passage.choices) && passage.choices.length > 0) {
    passage.choices.forEach((choice, idx) => {
      const choiceLabel = (choice && choice.label) ? String(choice.label) : `Choice ${idx + 1}`;
      const safeChoiceLabel = escapeEdgeLabel_(choiceLabel);

      const tgtQuest = choice && choice.target ? choice.target.quest : null;
      const tgtId = choice && choice.target ? String(choice.target.id) : null;

      const toEntry = resolveTargetEntry_(
        tgtQuest, tgtId,
        data, entryNodeByQuestNameAndId,
        missingTargetsByQuestNameAndId, questLinesBySlug, definedNodes, options
      );

      const parsedChoiceMacros = parseMacroBlocks_(String(choice.macros || ''));

      const chain = buildEdgeChainNodes_({
        basePrefix: `${makePassageNodeId_(pid)}_C_${idx + 1}`,
        questLines,
        definedNodes,
        ifs: parsedChoiceMacros.ifs,
        sets: parsedChoiceMacros.sets
      });

      if (chain.first) {
        edgeLines.push(`${fromExit} -->|${safeChoiceLabel}| ${chain.first}`);
        if (chain.last) edgeLines.push(`${chain.last} --> ${toEntry}`);
      } else {
        edgeLines.push(`${fromExit} -->|${safeChoiceLabel}| ${toEntry}`);
      }
    });
  }
}

function manageMacroNodes_(passage, baseId, questLines, definedNodes, edgeLines, entryNodeByQuestNameAndId, passageKey, exitNodeByQuestNameAndId) {
  const macros = String(passage.macros || '');
  const parsed = parseMacroBlocks_(macros);

  // IF nodes
  const ifConds = parsed.ifs;
  if (ifConds.length > 0) {
    let firstIfId = null;
    let prevIfId = null;

    ifConds.forEach((cond, idx) => {
      const ifId = `${baseId}_IF_${idx + 1}`;
      defineNodeApparence_(questLines, definedNodes, ifId, 'diamond', `if ${cond}`);
      if (!firstIfId) firstIfId = ifId;
      if (prevIfId) edgeLines.push(`${prevIfId} --> ${ifId}`);
      prevIfId = ifId;
    });

    edgeLines.push(`${prevIfId} --> ${baseId}`);
    entryNodeByQuestNameAndId.set(passageKey, firstIfId);
  } else {
    entryNodeByQuestNameAndId.set(passageKey, baseId);
  }

  //SET nodes
  const setOps = parsed.sets;
  if (setOps.length > 0) {
    let cur = baseId;
    setOps.forEach((op, idx) => {
      const setId = `${baseId}_SET_${idx + 1}`;
      defineNodeApparence_(questLines, definedNodes, setId, 'stored-data', `set: ${op}`);
      edgeLines.push(`${cur} --> ${setId}`);
      cur = setId;
    });
    exitNodeByQuestNameAndId.set(passageKey, cur);
  } else {
    exitNodeByQuestNameAndId.set(passageKey, baseId);
  }
}

function createPassageNode_(passage, pid, options, questLines, definedNodes, baseId, classLines) {
  const hasChoices = Array.isArray(passage.choices) && passage.choices.length > 0;
  const character = (passage.character || '').trim();
  const charUpper = character.toUpperCase();

  addCharacterClasses_(character, classLines, baseId);

  let shape = 'rect';
  if (hasChoices) shape = 'diamond';
  else if (charUpper === 'DOCUMENT') shape = 'doc';

  const label = buildPassageLabel_(character, passage.dialog, options.maxDialogChars);
  defineNodeApparence_(questLines, definedNodes, baseId, shape, label);
}

function addCharacterClasses_(character, classLines, baseId) {
  if (character === '') {    
    classLines.push(`class ${baseId} narrator;`);
  }
}

function prepareQuestMetadata_(data, questInfoByName, questLinesBySlug) {
  const questNames = Object.keys(data.quests);
  questNames.forEach((qName) => {
    const title = (qName && String(qName).trim() !== '') ? String(qName) : '(unnamed quest)';
    const slug = slugWithHash_(title);
    const subgraphId = `q_${slug}`;

    questInfoByName.set(qName, {
      slug,
      subgraphId,
      title
    });
    questLinesBySlug.set(slug, []);
  });
  return questNames;
}

function buildEdgeChainNodes_(params) {
  const basePrefix = params.basePrefix;
  const questLines = params.questLines;
  const definedNodes = params.definedNodes;

  const ifs = params.ifs || [];
  const sets = params.sets || [];

  let first = null;
  let prev = null;

  // IF diamonds first
  ifs.forEach((cond, idx) => {
    const id = `${basePrefix}_IF_${idx + 1}`;
    defineNodeApparence_(questLines, definedNodes, id, 'diamond', `if ${cond}`);
    if (!first) first = id;
    prev = id;
  });

  // SET stored-data nodes after
  sets.forEach((op, idx) => {
    const id = `${basePrefix}_SET_${idx + 1}`;
    defineNodeApparence_(questLines, definedNodes, id, 'stored-data', `set: ${op}`);
    if (!first) first = id;
    prev = id;
  });

  return {
    first: first,
    last: prev
  };
}

function resolveTargetEntry_(tgtQuest, tgtId, data, entryNodeByPassage, missingTargets, questLinesBySlug, definedNodes, options) {
  const key = makePassageKey_(tgtQuest, tgtId);

  const questObj = data.quests ? data.quests[tgtQuest] : null;
  const passageExists = questObj && questObj.passages && Object.prototype.hasOwnProperty.call(questObj.passages, tgtId);

  if (passageExists) {
    return entryNodeByPassage.get(key);
  }

  if (!options.includeMissingTargets) {
    return 'END';
  }

  if (!questLinesBySlug.has('__missing')) questLinesBySlug.set('__missing', []);
  const missingLines = questLinesBySlug.get('__missing');

  const missKey = `${String(tgtQuest)}|${String(tgtId)}`;
  if (!missingTargets.has(missKey)) {
    const missId = `MISSING_${hash4_(missKey)}`;
    missingTargets.set(missKey, missId);
    defineNodeApparence_(missingLines, definedNodes, missId, 'rect', `MISSING: ${tgtQuest}/${tgtId}`);
  }
  return missingTargets.get(missKey);
}

function defineNodeApparence_(questLines, definedNodes, nodeId, shape, label) {
  if (definedNodes.has(nodeId)) return;
  if (shape == 'rect' && (label === null || label.length === 0)) return;
  definedNodes.add(nodeId);

  //A@{ shape: rect, label: "A" }
  const safeLabel = escapeNodeLabel_(label);
  questLines.push(`${nodeId}@{ shape: ${shape}, label: "${safeLabel}" }`);
}

function makePassageNodeId_(passageId) {
  return `P_${getSafeIdPart_(passageId)}`;
}

function makePassageKey_(qName, pid) {
  return `${String(qName)}|${String(pid)}`;
}

function buildPassageLabel_(character, dialog, maxChars) {
  const c = (character && character.trim()) ? character.trim() : '';
  const cleaned = cleanText_(dialog || '');
  const snippet = (cleaned.length > maxChars) ? cleaned.slice(0, maxChars - 1) + '…' : cleaned;

  if (c) return `<i>${c}</i><br/>${snippet}`;
  return `${snippet}`;
}

function parseMacroBlocks_(macrosText) {
  const text = String(macrosText || '');
  const result = {
    ifs: [],
    sets: []
  };

  if (!text) return result;

  const starts = [];
  const re = /\(\s*(if:|set:|set\b)/gi;
  let m;
  while ((m = re.exec(text)) !== null) {
    starts.push(m.index);
  }

  starts.forEach((startIdx) => {
    const block = recoursivelyManageParentesis_(text, startIdx);
    if (!block) return;

    const inner = block.slice(1, -1).trim();
    const lower = inner.toLowerCase();

    if (lower.startsWith('if:')) {
      const cond = inner.slice(3).trim();
      if (cond) result.ifs.push(cond);
    } else if (lower.startsWith('set:')) {
      const op = inner.slice(4).trim();
      if (op) result.sets.push(op);
    } else if (lower.startsWith('set ')) {
      const op = inner.slice(3).trim();
      if (op) result.sets.push(op);
    }
  });

  return result;
}

function recoursivelyManageParentesis_(text, startIdx) {
  if (text[startIdx] !== '(') return null;
  let depth = 0;
  for (let i = startIdx; i < text.length; i++) {
    const ch = text[i];
    if (ch === '(') depth++;
    else if (ch === ')') depth--;
    if (depth === 0) return text.slice(startIdx, i + 1);
  }
  return null;
}

function cleanText_(s) {
  let t = String(s)
    .replace(/<\s*br\s*\/?\s*>/gi, '\n')
    .replace(/<\/\s*br\s*>/gi, '\n')
    .replace(/<\/\s*div\s*>/gi, '\n')
    .replace(/<\/\s*p\s*>/gi, '\n');

  t = t.replace(/<[^>]*>/g, '');

  t = t.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  t = t.replace(/\n{3,}/g, '\n\n');
  t = t.replace(/[ \t]{2,}/g, ' ').trim();

  return t.replace(/\n/g, '<br/>');
}

function escapeNodeLabel_(s) {
  return String(s)
    .replace(/\\/g, '\\\\')
    .replace(/"/g, "'")
    .replace(/\|/g, '¦');
}

function escapeEdgeLabel_(s) {
  return String(s).replace(/\|/g, '¦').replace(/\n/g, ' ').trim();
}

function getSafeIdPart_(s) {
  return String(s).replace(/[^A-Za-z0-9_]/g, '_');
}

function slugWithHash_(title) {
  const slug = String(title)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 24) || 'quest';
  return `${slug}_${hash4_(title)}`;
}

function hash4_(s) {
  let h = 5381;
  const str = String(s);
  for (let i = 0; i < str.length; i++) h = ((h << 5) + h) + str.charCodeAt(i);
  return (h >>> 0).toString(16).slice(-4);
}