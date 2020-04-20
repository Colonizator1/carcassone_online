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
  const playerCount = board.getRange('G1').getValue();
  const startCardId = '24';

  const color = ['green','red','blue','brown','yellow','black'];

function startGame() {
  board.getRange(2,1,100,100).clearContent();
  sheetGame.getRange(2,1,100,3).clearContent();
  deletImages(board);
  
  deckRange.copyTo(activeDeckRange);
  let shuffledArr = shuffle(activeDeckRange.getValues());
  activeDeckRange.setValues(shuffledArr);
  
  sheetGame.getRange(currentImageId).setValue(startCardId);
  sheetGame.getRange(currentRotate).setValue(1);
  insertPeoplesImage(playerCount);
}

function insertPeoplesImage (playerCount) {
  while (playerCount > 0) {
    let soldiersCount = 7;
    let drawnRaw = playerCount * 2;
    const horizontOffset = 43;
    let currentHorizontOffset = 0;
    let verticalOffset = [60,60,60,15,15,15,15];
    while (soldiersCount > 0) {
      if(soldiersCount == 3) {currentHorizontOffset = 0;}
      board.insertImage('https://classicsoffroad.com/carcassonne/img/peoples/knight-'+color[playerCount - 1]+'.png', 1, drawnRaw, currentHorizontOffset, verticalOffset[soldiersCount - 1]);
      currentHorizontOffset = currentHorizontOffset + horizontOffset;
      soldiersCount--;
    }
    playerCount--;
  }
}
function getCurrentRotate() {
  return sheetGame.getRange(currentRotate).getValue();
}
function getCurrentImageId() {
  return sheetGame.getRange(currentImageId).getValue();
}
function setRotate() {
  let newRotate = turnRotate(getCurrentRotate(), 4);
  sheetGame.getRange(currentRotate).setValue(newRotate);
}
function turnRotate(num, combinationCount) {
  return num >= combinationCount ? 1 : num + 1;
}

function move () {
  setCardInActiveCell();
  writeMoveStatistic();
  getRandomCart(activeDeckRange);
}
function setCardInActiveCell() {
  let activeCell = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  let activeRow = activeCell.getRow();
  let activeColumn = activeCell.getColumn();
  var imageUrl = 'http://classicsoffroad.com/carcassonne/img/' + getCurrentImageId() + '-' + getCurrentRotate() + '.jpg';
  if (activeCell.getFormula()) {
    ui.alert('Cell is full. Please check another one!');
  } else {
    activeCell.getCell(1, 1).setValue('=IMAGE("' + imageUrl + '")');
  }
}
function writeMoveStatistic() {
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
function getRandomInt(max) {
  return Math.floor(Math.random() * Math.floor(max));
}
function getRandomCart(someRange) {
  let lengthOfRange = getLengthRange(someRange);
  if (lengthOfRange == 0) {
      sheetGame.getRange(currentRotate).setValue('over');
      sheetGame.getRange(currentImageId).setValue('game');
      throw 'End of Game';
  }
  let column = someRange.getColumn();
  let actualDeck = sheetGame.getRange(2,column,lengthOfRange,column);
  let indexOfId = getRandomInt(lengthOfRange - 1) + 1;
  let randomCard = someRange.getCell(indexOfId,1).getValue();
  someRange.getCell(indexOfId,1).deleteCells(SpreadsheetApp.Dimension.ROWS);
  
  sheetGame.getRange(currentRotate).setValue(1);
  sheetGame.getRange(currentImageId).setValue(randomCard);
}
  
function shuffle(a) {
  let i = a.length - 1;
  while ( i > 0) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
    i--;
  }
  return a;
}
function getLengthRange(range) {
  let array = range.getValues();
  return array.filter(String).length;
}
function deletImages(range) {
  let images = range.getImages();
  images.map(function(img){img.remove();});
}
