const onOpen = () =>
  SpreadsheetApp.getUi().createMenu('RENOMATE').addItem('Importera från Trello', 'importFromTrello').addToUi();

const getBoardWithNestedResources = (boardId: string): any => {
  const params = {
    cards: 'all',
    card_pluginData: true,
    lists: 'all',
    members: 'all',
    customFields: true,
    labels: 'all',
    card_customFieldItems: true,
  };
  const paramsString = Object.entries(params)
    .map(([key, value]) => `${key}=${value}`)
    .join('&');
  const url = `https://api.trello.com/1/boards/${boardId}?${AUTH_PARAMS}&${paramsString}`;

  const response = UrlFetchApp.fetch(url);
  const json = response.getContentText();
  return JSON.parse(json);
};

const getBoardActions = (boardId: string): any[] => {
  const limit = 1000; // Maximum actions per page
  const baseUrl = `https://api.trello.com/1/boards/${boardId}/actions?${AUTH_PARAMS}&limit=${limit}`;
  const actions: any[] = [];
  let lastActionId: string | null = null;

  while (true) {
    const url = lastActionId ? `${baseUrl}&before=${lastActionId}` : baseUrl;
    const response = UrlFetchApp.fetch(url);
    const json = response.getContentText();
    const batchActions = JSON.parse(json);

    if (batchActions.length === 0) break; // No more actions to fetch
    actions.push(...batchActions);
    if (batchActions.length < limit) break; // Last batch fetched
    lastActionId = batchActions[batchActions.length - 1].id;
  }
  return actions;
};

const flattenObject = (obj: Record<string, any>, prefix: string = '', result: Record<string, any> = {}) => {
  Object.keys(obj).forEach((key) => {
    const newKey = prefix ? `${prefix}_${key}` : key;
    if (typeof obj[key] === 'object' && obj[key] !== null) {
      flattenObject(obj[key], newKey, result);
    } else {
      result[newKey] = obj[key];
    }
  });
  return result;
};

const createTableFromObjectsWithKeys = (objects: Record<string, any>[]): any[][] => {
  if (objects.length === 0) {
    return [];
  }

  const flattenedObjects = objects.map((object) => flattenObject(object));
  const allKeys = new Set(flattenedObjects.flatMap((object) => Object.keys(object)));
  const headers = [...allKeys];
  const table: any[][] = [headers];

  flattenedObjects.forEach((object) => {
    const row = headers.map((header) => object[header] ?? null);
    table.push(row);
  });

  return table;
};

const importFromTrello = (): void => {
  const board = getBoardWithNestedResources(TRELLO_BOARD_ID);
  console.log('got the board');
  const actions = getBoardActions(TRELLO_BOARD_ID);
  console.log('got the actions');
  const { labels, lists, members, customFields } = board;
  const cards = board.cards.map((card: any) => {
    card.customFieldItems.forEach((element: any) => {
      card[element.idCustomField] = element.value;
    });
    delete card.customFieldItems;
    return card;
  });
  const clearResource = {
    ranges: [LISTS_IMPORT_SHEET_NAME, CARDS_IMPORT_SHEET_NAME, ACTIONS_IMPORT_SHEET_NAME],
  };
  const updateResource = {
    valueInputOption: 'USER_ENTERED',
    includeValuesInResponse: false,
    data: [
      {
        range: CARDS_IMPORT_SHEET_NAME,
        values: createTableFromObjectsWithKeys(cards),
      },
      {
        range: LABELS_IMPORT_SHEET_NAME,
        values: createTableFromObjectsWithKeys(labels),
      },
      {
        range: LISTS_IMPORT_SHEET_NAME,
        values: createTableFromObjectsWithKeys(lists),
      },
      {
        range: MEMBERS_IMPORT_SHEET_NAME,
        values: createTableFromObjectsWithKeys(members),
      },
      {
        range: ACTIONS_IMPORT_SHEET_NAME,
        values: createTableFromObjectsWithKeys(actions),
      },
      {
        range: CUSTOMFIELDS_IMPORT_SHEET_NAME,
        values: createTableFromObjectsWithKeys(customFields),
      },
    ],
  };
  console.log('importing...');
  Sheets.Spreadsheets?.Values?.batchClear(clearResource, SPREADSHEET_ID);
  Sheets.Spreadsheets?.Values?.batchUpdate(updateResource, SPREADSHEET_ID);
};
