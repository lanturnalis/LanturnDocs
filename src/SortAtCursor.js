
function sortAtCursor() {
  var cursor = DocumentApp.getActiveDocument().getCursor();

  if (cursor && cursor.getElement()) {
    var exampleElement = cursor.getElement();

    while (!isSortableElement(exampleElement)) {
      // Bubble up to the parent looking for a sortable type
      exampleElement = exampleElement.getParent();
    }
    
    if (exampleElement.getType() == DocumentApp.ElementType.PARAGRAPH) {
      var paragraph = exampleElement.asParagraph();
      var headingLevel = getHeadingLevel(paragraph);
      if (headingLevel < 7) {
        return sortHeading(paragraph, headingLevel);
      }
    }
    else if (exampleElement.getType() == DocumentApp.ElementType.LIST_ITEM) {
      return sortListItem(exampleElement.asListItem());
    }
    
    DocumentApp.getUi().alert('No sortable element found');
  }
}

function sortHeading(headingToSort, headingLevelToSort) {
  var body = DocumentApp.getActiveDocument().getBody();
  var parent = headingToSort.getParent();

  // Find starting element of sort range
  var currentElement = headingToSort;
  var currentHeadingLevel = headingLevelToSort;
  while(currentHeadingLevel >= headingLevelToSort && currentElement.getPreviousSibling() != null) {
    currentElement = currentElement.getPreviousSibling();
    if (currentElement.getType() == DocumentApp.ElementType.PARAGRAPH) {
      currentHeadingLevel = getHeadingLevel(currentElement.asParagraph());
    }
  }
  while(currentHeadingLevel != headingLevelToSort) {
    currentElement = currentElement.getNextSibling();
    if (currentElement.getType() == DocumentApp.ElementType.PARAGRAPH) {
      currentHeadingLevel = getHeadingLevel(currentElement.asParagraph());
    }
  }
  var startingElement = currentElement;
  var startingIndex = parent.getChildIndex(startingElement)

  // Find stopping element of sort range
  currentElement = headingToSort;
  currentHeadingLevel = headingLevelToSort;
  var stoppingIndex = parent.getChildIndex(currentElement);
  while(currentElement != null && currentHeadingLevel >= headingLevelToSort) {
    currentElement = currentElement.getNextSibling();
    stoppingIndex++;
    if (currentElement != null && currentElement.getType() == DocumentApp.ElementType.PARAGRAPH) {
      currentHeadingLevel = getHeadingLevel(currentElement.asParagraph());
    }
  }
  while (stoppingIndex > startingIndex) {
    currentElement = parent.getChild(stoppingIndex-1);
    if (currentElement != null && currentElement.getType() == DocumentApp.ElementType.PARAGRAPH) {
      if (currentElement.asParagraph().getText() === "") {
        stoppingIndex--;
        continue;
      }
    }
    break;
  }

  // Collect all elements into heading buckets in sort range
  currentElement = startingElement
  var currentHeadingIndex = 0
  let headingMap = {};
  let headings = [];
  let headingIndexes = [];

  while(currentElement != null && parent.getChildIndex(currentElement) < stoppingIndex) {
    if (currentElement.getType() == DocumentApp.ElementType.PARAGRAPH) {
      
      let paragraph = currentElement.asParagraph();
      let headingLevel = getHeadingLevel(paragraph);

      if (headingLevel == headingLevelToSort) {
        headings.push(paragraph);
        currentHeadingIndex = headings.length - 1;
        headingIndexes.push(currentHeadingIndex);
        headingMap[currentHeadingIndex] = [currentElement];
        currentElement = currentElement.getNextSibling();
        continue;
      }
    }
    headingMap[currentHeadingIndex].push(currentElement);
    currentElement = currentElement.getNextSibling();
  }

  // Remove all elements in sort range
  currentElement = startingElement
  for (let index = stoppingIndex - 1; index >= startingIndex; index--) {
    if (parent.getChild(index).isAtDocumentEnd()) {
      parent.appendParagraph("")
    }
    parent.removeChild(parent.getChild(index));
  }

  // Sort headings
  headingIndexes = headingIndexes.sort((a, b) => {
    let aText = headings[a].getText();
    let bText = headings[b].getText();
    let compareResult = aText.localeCompare(bText);
    return compareResult;
  });

  var index = startingIndex;
  for (let i = 0; i < headingIndexes.length; i++) {
    let headingIndex = headingIndexes[i];
    var bucketElements = headingMap[headingIndex];
    
    for (let j = 0; j < bucketElements.length; j++) {
      let element = bucketElements[j];
      insertElement(parent, index, element);
      index++;
    }
  }
}

function sortListItem(listItemToSort) {
var body = DocumentApp.getActiveDocument().getBody();
  let parent = listItemToSort.getParent();
  let nestingLevelToSort = listItemToSort.getNestingLevel();

  // Find starting element of sort range
  var currentElement = listItemToSort;
  var currentNestingLevel = nestingLevelToSort;
  while(currentElement.getPreviousSibling() != null && currentElement.getType() == DocumentApp.ElementType.LIST_ITEM && currentNestingLevel >= nestingLevelToSort) {
    currentElement = currentElement.getPreviousSibling();
    if (currentElement.getType() == DocumentApp.ElementType.LIST_ITEM) {
      currentNestingLevel = currentElement.asListItem().getNestingLevel();
    } else {
      currentNestingLevel = -1;
    }
  }
  while(currentNestingLevel != nestingLevelToSort) {
    currentElement = currentElement.getNextSibling();
    if (currentElement.getType() == DocumentApp.ElementType.LIST_ITEM) {
      currentNestingLevel = currentElement.asListItem().getNestingLevel();
    }
  }
  var startingElement = currentElement;
  var startingIndex = parent.getChildIndex(startingElement)

  // Find stopping element of sort range
  currentElement = listItemToSort;
  currentNestingLevel = nestingLevelToSort;
  var stoppingIndex = parent.getChildIndex(currentElement);
  while(currentElement != null && currentElement.getType() == DocumentApp.ElementType.LIST_ITEM && currentNestingLevel >= nestingLevelToSort) {
    currentElement = currentElement.getNextSibling();
    stoppingIndex++;
    if (currentElement != null && currentElement.getType() == DocumentApp.ElementType.LIST_ITEM) {
      currentNestingLevel = currentElement.asListItem().getNestingLevel();
    }
  }

  // Collect all elements into listItem buckets in sort range
  currentElement = startingElement
  var currentListItemIndex = 0
  let listItemMap = {};
  let listItems = [];
  let listItemIndexes = [];

  while(currentElement != null && parent.getChildIndex(currentElement) < stoppingIndex) {

    if (currentElement.getType() == DocumentApp.ElementType.LIST_ITEM) {
      
      let listItem = currentElement.asListItem();
      let listItemLevel = listItem.getNestingLevel();

      if (listItemLevel == nestingLevelToSort) {
        listItems.push(listItem);
        currentListItemIndex = listItems.length - 1;
        listItemIndexes.push(currentListItemIndex);
        listItemMap[currentListItemIndex] = [currentElement];
        currentElement = currentElement.getNextSibling();
        continue;
      }
    }
    listItemMap[currentListItemIndex].push(currentElement);
    currentElement = currentElement.getNextSibling();
  }

  // Remove all elements in sort range
  currentElement = startingElement
  for (let index = stoppingIndex - 1; index >= startingIndex; index--) {
    if (parent.getChild(index).isAtDocumentEnd()) {
      parent.appendParagraph("")
    }
    parent.removeChild(parent.getChild(index));
  }

  // Sort listItems
  listItemIndexes = listItemIndexes.sort((a, b) => {
    let aText = listItems[a].getText();
    let bText = listItems[b].getText();
    let compareResult = aText.localeCompare(bText);
    return compareResult;
  });

  var index = startingIndex;
  for (let i = 0; i < listItemIndexes.length; i++) {
    let listItemIndex = listItemIndexes[i];
    var bucketElements = listItemMap[listItemIndex];
    
    for (let j = 0; j < bucketElements.length; j++) {
      let element = bucketElements[j];
      insertElement(parent, index, element);
      index++;
    }
  }
}

function insertElement(parent, index, element) {
  switch(element.getType()) {
  case DocumentApp.ElementType.PARAGRAPH:
    parent.insertParagraph(index, element.asParagraph());
    break;
  case DocumentApp.ElementType.TABLE:
    parent.insertTable(index, element.asTable());
    break;
  case DocumentApp.ElementType.LIST_ITEM:
    parent.insertListItem(index, element.asListItem());
    break;
  }
}

function isSortableElement(element) {
  switch(element.getType()) {
  case DocumentApp.ElementType.DOCUMENT:
  case DocumentApp.ElementType.LIST_ITEM:
  case DocumentApp.ElementType.PARAGRAPH:
    return true;
  default:
    return false;
  }
}

function getHeadingLevel(paragraph) {
  switch(paragraph.getHeading()) {
  case DocumentApp.ParagraphHeading.HEADING1: return 1;
  case DocumentApp.ParagraphHeading.HEADING2: return 2;
  case DocumentApp.ParagraphHeading.HEADING3: return 3;
  case DocumentApp.ParagraphHeading.HEADING4: return 4;
  case DocumentApp.ParagraphHeading.HEADING5: return 5;
  case DocumentApp.ParagraphHeading.HEADING6: return 6;
  default:
    return 7;
  }
}
