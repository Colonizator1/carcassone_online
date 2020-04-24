  const SPREADSHEET_URL = 'PASTE_YOUR_SPREADSHEET_URL';

  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const ui = SpreadsheetApp.getUi();
  const board = ss.getSheetByName('karkasson');
  const sheetGame = ss.getSheetByName('game_process');
  const sheetDeck = ss.getSheetByName('deck');
  const deckRange = sheetDeck.getRange('D2:D72');
  const activeDeckRange = sheetGame.getRange('F2:F72');
  const currentRotate = 'E2';
  const currentImageId = 'D2';
  const currentPlayer = sheetGame.getRange('G2');
  const playerCount = board.getRange('G1').getValue();
  const activePlayer = sheetGame.getRange('H2');
  const startCardId = '24';

  const color = ['green','red','blue','brown','yellow','black'];

const startGame = () => {
  let response = ui.alert('Are you sure you want to start new game?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
	  board.getRange(2,1,100,100).clearContent();
	  sheetGame.getRange(2,1,100,3).clearContent();
	  deletImages(board);
	  
	  deckRange.copyTo(activeDeckRange);
	  let shuffledArr = shuffle(activeDeckRange.getValues());
	  activeDeckRange.setValues(shuffledArr);
	  
	  sheetGame.getRange(currentImageId).setValue(startCardId);
	  sheetGame.getRange(currentRotate).setValue(1);
	  currentPlayer.setValue(0);
	  activePlayer.setValue(color[0]);
	  insertPeoplesImage(playerCount);
  } else {
    return false;
  }
}
const insertPeoplesImage = (playerCount) => {
  while (playerCount > 0) {
    let soldiersCount = 7;
    let drawnRaw = playerCount * 2;
    const horizontOffset = 43;
    let currentHorizontOffset = 0;
    let verticalOffset = [60,60,60,15,15,15,15];
    while (soldiersCount > 0) {
      if(soldiersCount == 3) {currentHorizontOffset = 0;}
      board.insertImage('https://raw.githubusercontent.com/Colonizator1/carcassone_online/master/img/peoples/knight-'+color[playerCount - 1]+'.png', 1, drawnRaw, currentHorizontOffset, verticalOffset[soldiersCount - 1]);
      currentHorizontOffset = currentHorizontOffset + horizontOffset;
      soldiersCount--;
    }
    playerCount--;
  }
}
const makeMove = () => {
  if(setCardInActiveCell()) {
    setNextPlayer(playerCount);
    writeMoveStatistic();
    let randomCard = getRandomCart(activeDeckRange);
    setRandomCard(randomCard);
    randomCard.deleteCells(SpreadsheetApp.Dimension.ROWS);
  };
}
const setCardInActiveCell = () => {
  let activeCell = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  let activeRow = activeCell.getRow();
  let activeColumn = activeCell.getColumn();
  var imageUrl = 'https://raw.githubusercontent.com/Colonizator1/carcassone_online/master/img/' + getCurrentImageId() + '-' + getCurrentRotate() + '.jpg';
  if (activeCell.getFormula()) {
    ui.alert('Cell is full. Please check another one!');
    return false;
  } else {
    activeCell.getCell(1, 1).setValue('=IMAGE("' + imageUrl + '")');
    return true;
  }
}
const setNextPlayer = (players) => {
  let newPlayer = turnRotate(currentPlayer.getValue(), players - 1);
  currentPlayer.setValue(newPlayer);
  activePlayer.setValue(color[newPlayer]);
}
const writeMoveStatistic = () => {
  let moveRange = sheetGame.getRange('A2:C100');
  var i = 1;
  while (i < 100) {
    let lastFullMoveId = moveRange.getCell(i, 1).getValue();
    let savedImageRange = moveRange.getCell(i, 2).getValue();
    let savedRotation = moveRange.getCell(i, 3).getValue();
    if (lastFullMoveId) {
      i++;
    } else {
      moveRange.getCell(i, 1).setValue(i);
      moveRange.getCell(i, 2).setValue(getCurrentImageId());
      moveRange.getCell(i, 3).setValue(getCurrentRotate());
      break;
    }
  }
}
const setRandomCard = (card) => {
  sheetGame.getRange(currentImageId).setValue(card.getValue());
}
const getRandomCart = (someRange) => {
  let lengthOfRange = getLengthRange(someRange);
  if (lengthOfRange == 0) {
      sheetGame.getRange(currentRotate).setValue('over');
      sheetGame.getRange(currentImageId).setValue('game');
      throw 'End of Game';
  }
  let column = someRange.getColumn();
  let actualDeck = sheetGame.getRange(2,column,lengthOfRange,column);
  let indexOfId = getRandomInt(lengthOfRange - 1) + 1;
  let randomCard = someRange.getCell(indexOfId,1);
  return randomCard;
}
const getRandomInt = (max) => {
  return Math.floor(Math.random() * Math.floor(max));
}

const shuffle = (a) => {
  let i = a.length - 1;
  while ( i > 0) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
    i--;
  }
  return a;
}
const getLengthRange = (range) => {
  let array = range.getValues();
  return array.filter(String).length;
}
const deletImages = (range) => {
  let images = range.getImages();
  images.map(function(img){img.remove();});
}

const getCurrentRotate = () => {
  return sheetGame.getRange(currentRotate).getValue();
}
const getCurrentImageId = () => {
  return sheetGame.getRange(currentImageId).getValue();
}
const setRotate = () => {
  let newRotate = turnRotate(getCurrentRotate(), 4, 1);
  sheetGame.getRange(currentRotate).setValue(newRotate);
}


const turnRotate = (currentNum, combinationCount, startNum = 0) => {
  return currentNum >= combinationCount ? startNum : currentNum + 1;
}
}
const makeMove = () => {
  if(setCardInActiveCell()) {
    setNextPlayer(playerCount);
    writeMoveStatistic();
    let randomCard = getRandomCart(activeDeckRange);
    setRandomCard(randomCard);
    randomCard.deleteCells(SpreadsheetApp.Dimension.ROWS);
  };
}
const setCardInActiveCell = () => {
  let activeCell = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  let activeRow = activeCell.getRow();
  let activeColumn = activeCell.getColumn();
  var imageUrl = 'http://classicsoffroad.com/carcassonne/img/' + getCurrentImageId() + '-' + getCurrentRotate() + '.jpg';
  if (activeCell.getFormula()) {
    ui.alert('Cell is full. Please check another one!');
    return false;
  } else {
    activeCell.getCell(1, 1).setValue('=IMAGE("' + imageUrl + '")');
    return true;
  }
}
const setNextPlayer = (players) => {
  let newPlayer = turnRotate(currentPlayer.getValue(), players - 1);
  currentPlayer.setValue(newPlayer);
  activePlayer.setValue(color[newPlayer]);
}
const writeMoveStatistic = () => {
  let moveRange = sheetGame.getRange('A2:C100');
  var i = 1;
  while (i < 100) {
    let lastFullMoveId = moveRange.getCell(i, 1).getValue();
    let savedImageRange = moveRange.getCell(i, 2).getValue();
    let savedRotation = moveRange.getCell(i, 3).getValue();
    if (lastFullMoveId) {
      i++;
    } else {
      moveRange.getCell(i, 1).setValue(i);
      moveRange.getCell(i, 2).setValue(getCurrentImageId());
      moveRange.getCell(i, 3).setValue(getCurrentRotate());
      break;
    }
  }
}
const setRandomCard = (card) => {
  sheetGame.getRange(currentImageId).setValue(card.getValue());
}
const getRandomCart = (someRange) => {
  let lengthOfRange = getLengthRange(someRange);
  if (lengthOfRange == 0) {
      sheetGame.getRange(currentRotate).setValue('over');
      sheetGame.getRange(currentImageId).setValue('game');
      throw 'End of Game';
  }
  let column = someRange.getColumn();
  let actualDeck = sheetGame.getRange(2,column,lengthOfRange,column);
  let indexOfId = getRandomInt(lengthOfRange - 1) + 1;
  let randomCard = someRange.getCell(indexOfId,1);
  return randomCard;
}
const getRandomInt = (max) => {
  return Math.floor(Math.random() * Math.floor(max));
}

const shuffle = (a) => {
  let i = a.length - 1;
  while ( i > 0) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
    i--;
  }
  return a;
}
const getLengthRange = (range) => {
  let array = range.getValues();
  return array.filter(String).length;
}
const deletImages = (range) => {
  let images = range.getImages();
  images.map(function(img){img.remove();});
}

const getCurrentRotate = () => {
  return sheetGame.getRange(currentRotate).getValue();
}
const getCurrentImageId = () => {
  return sheetGame.getRange(currentImageId).getValue();
}
const setRotate = () => {
  let newRotate = turnRotate(getCurrentRotate(), 4, 1);
  sheetGame.getRange(currentRotate).setValue(newRotate);
}


const turnRotate = (currentNum, combinationCount, startNum = 0) => {
  return currentNum >= combinationCount ? startNum : currentNum + 1;
}
