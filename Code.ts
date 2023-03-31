// Original Typescript code: https://github.com/vilhelm-k/Trello

interface TrelloList {
  id: string;
  name: string;
  closed: boolean;
  pos: number;
  softLimit: string;
  idBoard: string;
  subscribed: boolean;
  limits: any;
}

interface TrelloCard {
  id: string;
  address: string | null;
  badges: any;
  checkItemStates: string[];
  closed: boolean;
  coordinates: string | null;
  creationMethod: string | null;
  dateLastActivity: string;
  desc: string;
  descData: any;
  due: string | null;
  dueReminder: string | null;
  email: string;
  idBoard: string;
  idChecklists: Array<string | any>;
  idLabels: Array<string | any>;
  idList: string;
  idMembers: string[];
  idMembersVoted: string[];
  idShort: number;
  idAttachmentCover: string;
  labels: string[];
  limits: any;
  locationName: string | null;
  manualCoverAttachment: boolean;
  name: string;
  pos: number;
  shortLink: string;
  shortUrl: string;
  subscribed: boolean;
  url: string;
  cover: any;
}

interface TrelloAction {
  id: string;
  idMemberCreator: string;
  data: ActionData;
  type: string;
  date: string;
  limits: any;
  display: any;
  memberCreator: any;
}

interface ActionData {
  board?: {
    id: string;
    name: string;
    shortLink: string;
  };
  card?: {
    id: string;
    name: string;
    idShort: number;
    shortLink: string;
  };
  listBefore?: {
    id: string;
    name: string;
  };
  listAfter?: {
    id: string;
    name: string;
  };
  list?: {
    id: string;
    name: string;
  };
  old?: {
    idList?: string;
    name?: string;
    desc?: string;
    due?: string | null;
    closed?: boolean;
    pos?: number;
  };
}

const onOpen = () =>
  SpreadsheetApp.getUi().createMenu('RENOMATE').addItem('Importera från Trello', 'importFromTrello').addToUi();

const getTrelloCardsFromBoard = (boardId: string): TrelloCard[] => {
  const url = `${URL_BASE}/boards/${boardId}/cards?${AUTH_PARAMS}`;
  const response = UrlFetchApp.fetch(url);
  const json = response.getContentText();
  return JSON.parse(json);
};

const getTrelloListsFromBoard = (boardId: string): TrelloList[] => {
  const url = `${URL_BASE}/boards/${boardId}/lists?${AUTH_PARAMS}`;
  const response = UrlFetchApp.fetch(url);
  const json = response.getContentText();
  return JSON.parse(json);
};

const getTrelloCardActions = (cardId: string): TrelloAction[] => {
  const maxPage = 19;
  const actionsPerPage = 50;
  const allActions: TrelloAction[] = [];

  for (let page = 0; page <= maxPage; page++) {
    const url = `${URL_BASE}/cards/${cardId}/actions?filter=updateCard,createCard&page=${page}&${AUTH_PARAMS}`;
    const response = UrlFetchApp.fetch(url);
    const json = response.getContentText();
    const actions = JSON.parse(json);

    if (actions.length === 0) break;
    allActions.push(...actions);
    if (actions.length < actionsPerPage) break;
  }

  return allActions;
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
  const lists = getTrelloListsFromBoard(TRELLO_BOARD_ID);
  const cards = getTrelloCardsFromBoard(TRELLO_BOARD_ID);
  const cardMoveActions = cards
    .filter((card) => card.idList !== TRELLO_SPAWN_LIST)
    .flatMap((card) => getTrelloCardActions(card.id));

  const clearResource = {
    ranges: [LISTS_IMPORT_SHEET_NAME, CARDS_IMPORT_SHEET_NAME, CARD_ACTIONS_IMPORT_SHEET_NAME],
  };
  const updateResource = {
    valueInputOption: 'USER_ENTERED',
    includeValuesInResponse: false,
    data: [
      {
        range: LISTS_IMPORT_SHEET_NAME,
        values: createTableFromObjectsWithKeys(lists),
      },
      {
        range: CARDS_IMPORT_SHEET_NAME,
        values: createTableFromObjectsWithKeys(cards),
      },
      {
        range: CARD_ACTIONS_IMPORT_SHEET_NAME,
        values: createTableFromObjectsWithKeys(cardMoveActions),
      },
    ],
  };

  Sheets.Spreadsheets?.Values?.batchClear(clearResource, SPREADSHEET_ID);
  Sheets.Spreadsheets?.Values?.batchUpdate(updateResource, SPREADSHEET_ID);
};
