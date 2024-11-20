let itemsPerPageCount = 5;
let pageCount = 1;
let accessToken = '';
let selectedMenuItemId = '';
let menuTitles = {};
let accessTokenStatus = 'refreshed';

const expiresIn = 3599; // Example expiration time in seconds

const slideApiUrl = "https://www.googleapis.com/drive/v3/files?q=(mimeType='application/vnd.google-apps.presentation' or mimeType='application/vnd.openxmlformats-officedocument.presentationml.presentation') and  trashed=false&fields=files(id,name,webViewLink)";

const tokenEndpoint = 'https://oauth2.googleapis.com/token';

let user_signed_in = false;

function isEmpty(obj) {
  return obj === undefined || obj === null || Object.keys(obj).length === 0;
}

function is_user_signed_in() {
  return user_signed_in;
}

const constructAuthUrl = () => {
  const AUTH_URL =
    `https://accounts.google.com/o/oauth2/auth\
?client_id=${config.CLIENT_ID}\
&response_type=code\
&redirect_uri=${encodeURIComponent(config.REDIRECT_URL)}\
&include_granted_scopes=true
&scope=https://www.googleapis.com/auth/drive
&access_type=offline`;
  return AUTH_URL;
}

const exchangeAuthorizationCode = async (authorizationCode) => {
  const response = await fetch(tokenEndpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: new URLSearchParams({
      client_id: config.CLIENT_ID,
      client_secret: config.CLIENT_SECRET,
      code: authorizationCode,
      redirect_uri: config.REDIRECT_URL,
      grant_type: 'authorization_code',
    }),
  })
  const tokenData = await response.json();
  accessToken = tokenData.access_token;
  chrome.storage.local.set({ 'access_token_expires_at':  Date.now()+ (expiresIn*1000)}, function() {
  });
  if(tokenData?.refresh_token) {
    chrome.storage.local.set({ 'refresh_token': tokenData.refresh_token }, function() {
      // console.log('Refresh Token stored locally.');
    });
  }
  // console.log('accessToken: ', accessToken);
}

const getAuthorizationCode = async (url, paramName) => {
  const params = new URLSearchParams(url.split('?')[1]);
  return params.get(paramName);
}

/* Login the user */

const loginUser = (AUTH_URL) => {
  chrome.identity.launchWebAuthFlow({
    'url': AUTH_URL,
    'interactive': true
  }, async function (redirect_url) {
    // console.log('redirect_url: ', redirect_url);
    if (chrome.runtime.lastError) {
      console.log('Problem signing in');
    } else {
      // console.log('redirect_url: ', redirect_url);
      const authorizationCode = await getAuthorizationCode(redirect_url, 'code');
      await exchangeAuthorizationCode(authorizationCode);
      user_signed_in = true;
      createMenus();
      chrome.action.setPopup({ popup: './logout.html' }, () => {});
    }
  })
}

/* ---- End ---- */

const removeMenus = async() => {
  chrome.contextMenus.remove('newPresentation', function () {
    console.log("newPresentation menu removed");
  });
  chrome.contextMenus.remove('ExistingSlide', function () {
    console.log("ExistingSlide menu removed");
  });
  chrome.contextMenus.remove('Recent', function () {
    console.log("Recent menu removed");
  });
}

/* Logout the user */
const logoutUser = async() => {
  // Removing Menus
  await removeMenus();
  chrome.action.setPopup({ popup: './login.html' }, () => {
  });
}
/* ---- End ---- */


const refreshAccessToken = async () => {
  let refreshToken = '';
  await chrome.storage.local.get(['refresh_token'], async function (result) {
    refreshToken = result['refresh_token'];
    accessTokenStatus = 'refreshing';
    const response = await fetch(tokenEndpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      body: new URLSearchParams({
        grant_type: 'refresh_token',
        client_id: config.CLIENT_ID,
        client_secret: config.CLIENT_SECRET,
        redirect_uri: config.REDIRECT_URL,
        refresh_token: refreshToken
      })
    });
    if (!response.ok) {
      console.log('Failed to refresh access token');
    } else {
      const res = await response.json();
      accessToken = res.access_token;
      chrome.storage.local.set({ 'access_token_expires_at':  Date.now()+(expiresIn*1000)}, function() {
      });
      accessTokenStatus = 'refreshed';
    }
    return accessToken;
  });
}

/* Access token expiration check */

const checkAccessTokenExpiry = async () => {
  chrome.storage.local.get(['access_token_expires_at'], async function (result) {
    const expiresAt = result['access_token_expires_at'];
    if (Date.now() >= expiresAt || accessToken == '') {
      console.log('Expired');
      accessTokenStatus = 'expired';
      await refreshAccessToken();
    }
  });
  return true;
}
/* ---- End ---- */

/* Get all google slides list for the logged in account */
const getAllGoogleSlideList = async() => {
  const presList = [];
  menuTitles = {};

  const create = async() => {
    if(accessTokenStatus == 'refreshed') {
      await fetch(slideApiUrl, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
        },
      })
      .then(response => response.json())
      .then(data => {
        data?.files?.forEach((eachSlide) => {
          const obj = { id: eachSlide.id, name: eachSlide.name };
          menuTitles[eachSlide.id] = eachSlide.name;
          presList.push(obj);
        });
      })
      .catch(error => {
        console.error('Error:', error);
      });
      return presList;
    } else {
      await logoutUser();
    }
  }

  await checkAccessTokenExpiry();
  if(accessToken == '' || accessTokenStatus == 'refreshing' || accessTokenStatus == 'expired') {
    // setTimeout(() => {
    //   console.log("Delayed for 250 milli second in new presentation.");
      
    // }, "800");
    return await create();
  } else if(accessTokenStatus == 'refreshed') { 
    return create(); 
  }
 
}
/* ---- End ---- */

const createChildContextMenu = async (slideList, currentPage, itemsPerPage, isMenu = false) => {
  const startIndex = currentPage * itemsPerPageCount;
  const endIndex = startIndex + itemsPerPage;
  const visibleItems = slideList.slice(startIndex, endIndex);

  // Create context menu items based on visibleItems
  visibleItems.forEach((item, index) => {
    chrome.contextMenus.create({
      id: item.id,
      title: item.name,
      type: "radio",
      checked: false,
      parentId: "ExistingSlide",
      contexts: ["all"],
    });
  });

  // If there are more items, add "Show More" option
  if (endIndex < slideList.length) {
    chrome.contextMenus.create({
      id: "ShowMore",
      title: "Show More",
      parentId: "ExistingSlide",
      contexts: ["all"],
    });
  }
  if (isMenu) pageCount++;
}

const createMainMenu = async () => {
  chrome.contextMenus.create({
    id: "mainMenu",
    title: "SaveToGSlides",
    contexts: ["all"]
  });
};

const createMenus = async() => {
  // Create the main context menu item
  chrome.contextMenus.create({
    id: "newPresentation",
    title: "Add in new slide",
    parentId: "mainMenu",
    contexts: ["all"]
  });

  // Create submenus for the main menu item
  chrome.contextMenus.create({
    id: "ExistingSlide",
    title: "Choose Existing Slide",
    parentId: "mainMenu",
    contexts: ["all"]
  });

  chrome.contextMenus.create({
    id: "Recent",
    title: "Recent",
    parentId: "mainMenu",
    contexts: ["all"]
  });

  const presentationList = await getAllGoogleSlideList();
  await createChildContextMenu(presentationList, 0, itemsPerPageCount);
}

const handleClick = async (info, tab) => {
  // Update the selected menu item
  selectedMenuItemId = info.menuItemId;
  // Update the context menu appearance
  chrome.contextMenus.update("newPresentation", { "title": "Add in new slide" });
  chrome.contextMenus.remove('ExistingSlide', () => {
  });
  chrome.contextMenus.create({
    id: "ExistingSlide",
    title: "Choose Existing Slide",
    parentId: "mainMenu",
    contexts: ["all"]
  });
  const currPresentationList = await getAllGoogleSlideList();
  await createChildContextMenu(currPresentationList, 0, itemsPerPageCount);

  // Apply a tick mark to the selected menu item
  if (info.parentMenuItemId === "ExistingSlide") {
    chrome.storage.local.set({ 'Recent': selectedMenuItemId }, function() {
    });
    chrome.contextMenus.update('Recent', { "title": "Recent \u000A"+menuTitles[selectedMenuItemId] });
    chrome.contextMenus.update(selectedMenuItemId, { "checked": true });
  } else if(info.menuItemId === "newPresentation") {
    chrome.storage.local.set({ 'Recent': currPresentationList[0].id }, function() {
    });
    chrome.contextMenus.update('Recent', { "title": "Recent \u000A"+currPresentationList[0].name });
  }
}

const getSlideDetails = async(presentationId) => {
  const getSlideUrl = `https://slides.googleapis.com/v1/presentations/${presentationId}`

  let slideData = {};
  const create = async() => {
    if(accessTokenStatus == 'refreshed') {
      await fetch(getSlideUrl, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
        },
      })
      .then(response => response.json())
      .then(data => {
        slideData = data;
      })
      .catch(error => {
        console.log('Error:', error);
      });
      return slideData;
    } else {
      await logoutUser();
    }
  }

  await checkAccessTokenExpiry();
  if(accessToken == '' || accessTokenStatus == 'refreshing' || accessTokenStatus == 'expired') {
    // setTimeout(() => {
    //   console.log("Delayed for 800 milli second in new presentation.");
    // return create();
    // }, "800");
    return create();
  } else if(accessTokenStatus == 'refreshed') { 
    return create(); 
  }
  
}

const getSlideInsertionPositions = async (presentationId) => {
  let slideDetails = await getSlideDetails(presentationId);
  let slideEndIndex = undefined, slideStartIndex = undefined, slideLength, lastTextBoxId;
  if (!isEmpty(slideDetails) && 'slides' in slideDetails) { 
    slideLength = slideDetails.slides.length - 1;
    const pageElementsLength = slideDetails.slides[slideLength]?.pageElements?.length - 1 || undefined;
    if(pageElementsLength != undefined) {
      const lastTextElementLength = slideDetails.slides[slideLength].pageElements[pageElementsLength].shape?.text?.textElements?.length - 1;
      slideEndIndex = slideDetails.slides[slideLength].pageElements[pageElementsLength].shape?.text?.textElements[lastTextElementLength]?.endIndex;
      slideStartIndex = slideDetails.slides[slideLength].pageElements[pageElementsLength].shape?.text?.textElements[lastTextElementLength]?.startIndex;
      lastTextBoxId = slideDetails.slides[slideLength].pageElements[pageElementsLength]?.objectId;
    }
  }
  return { slideEndIndex, slideStartIndex, slideDetails, slideLength, lastTextBoxId};
}


/* Creates New slide in a Presentation */
const createNewSlide = async (presentationId, layout) => {
  const postSlideApiUrl = `https://slides.googleapis.com/v1/presentations/${presentationId}:batchUpdate`;
  let insertionResult;
  const create = async() => {
    if(accessTokenStatus == 'refreshed') {
      const { formattedSingleSlideName }  = getFormattedNames();
      await fetch(postSlideApiUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          requests: [
            {
              createSlide: {
                objectId: formattedSingleSlideName,
                slideLayoutReference: {
                  predefinedLayout: layout
                }
              }
            }
          ]
        })
      })
      .then(response => response.json())
      .then(async createSlideResponse => {
        insertionResult = await getSlideInsertionPositions(presentationId);
        // return insertionResult;
      })
      .catch(error => {
        console.error('Error:', error);
      });
    } else {
      await logoutUser();
    }
    return insertionResult;
  }

  await checkAccessTokenExpiry();
  if(accessToken == '' || accessTokenStatus == 'refreshing' || accessTokenStatus == 'expired') {
    // setTimeout(() => {
    //   console.log("Delayed for 250 milli second in new presentation.");
    //   const res = create();
    //   console.log('resss in create new slide1', res);
    //   return res;
    // }, "800");
    return create();
  } else if(accessTokenStatus == 'refreshed') { 
    const res = create();
    return res;
    // return create(); 
  }
  
};

const addBulletins = async(presentationId,lastTextBoxId)=>{
  const postSlideApiUrl = `https://slides.googleapis.com/v1/presentations/${presentationId}:batchUpdate`;
  const requests = [];
  const bulletinRequest =  {
    'createParagraphBullets':{
    'objectId': lastTextBoxId,
    'textRange':
      {
        'type':'ALL',
      },
      "bulletPreset": "BULLET_DISC_CIRCLE_SQUARE",
    }
  }
  requests.push(bulletinRequest);
  const create = async() => {
    if(accessTokenStatus == 'refreshed') {
      await fetch(postSlideApiUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({"requests":requests})
  
      })
        .then(response => response.json())
        .catch(error => {
          console.error('Error:', error);
        });
    } else {
      await logoutUser();
    }
  }

  await checkAccessTokenExpiry();
  if(accessToken == ''|| accessTokenStatus == 'refreshing' || accessTokenStatus == 'expired') {
    // setTimeout(() => {
    //   console.log("Delayed for 250 milli second in new presentation.");
    //   create();
    // }, "800");
    return create();
  } else if(accessTokenStatus == 'refreshed') { 
    create(); 
  }
  
}

const insertTextToLastSlide = async (presentationId, lastTextBoxId, info, slideStartIndex) => {
  const postSlideApiUrl = `https://slides.googleapis.com/v1/presentations/${presentationId}:batchUpdate`;
  let textmsg = '';
  const requests = [];
  /* For text */
  if (info?.selectionText) {
    textmsg += info.selectionText;
    requests.push(
      {
        "insertText": {
          "objectId": lastTextBoxId,
          "text": `${textmsg}\n`,
          "insertionIndex": slideStartIndex
        }
      }
    );
  }

  const create = async() => {
    if(accessTokenStatus == 'refreshed') {
      await fetch(postSlideApiUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({"requests":requests}) 
      })
        .then(response => response.json())
        .catch(error => {
          console.error('Error:', error);
        });
        // Add bulletins after text is inserted, they cannot be inserted during text insertion
        // Need to be inserted as seperate call
        if(info?.selectionText){
          addBulletins(presentationId,lastTextBoxId);
        }
    } else {
      await logoutUser();
    }
  }

  await checkAccessTokenExpiry();
  if(accessToken == '' || accessTokenStatus == 'refreshing' || accessTokenStatus == 'expired') {
    // setTimeout(() => {
    //   console.log("Delayed for 250 milli second in new presentation.");
    //   create();
    // }, "800");
    return create();
  } else if(accessTokenStatus == 'refreshed') { 
    create(); 
  }
}

const insertImageToLastSlide = async (presentationId, info, slideDetails) => {
  const postSlideApiUrl = `https://slides.googleapis.com/v1/presentations/${presentationId}:batchUpdate`;
  const requests = [];
  /* For image */
  if (info?.mediaType && info.mediaType === "image") {
    const imageUrl = info.srcUrl
    const allSlides = slideDetails?.slides;
    const slideLength = allSlides.length;
    const lastSlide = allSlides[slideLength-1];
    const pageId = lastSlide.objectId;
    requests.push({
      "createImage": {
        "url": imageUrl,
        "elementProperties": {
          "pageObjectId": pageId,
          "size": {
            "height": {
              "magnitude": 300,
              "unit": "PT"
            },
            "width": {
              "magnitude": 400,
              "unit": "PT"
            }
          },
          "transform": {
            "scaleX": 1,
            "scaleY": 1,
            "translateX": 100,
            "translateY": 100,
            "unit": "PT"
          }
        }
      }
    });
  }

  const create = async() => {
    if(accessTokenStatus == 'refreshed') {
      await fetch(postSlideApiUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({"requests":requests})
  
      })
      .then(response => response.json())
      .catch(error => {
        console.error('Error:', error);
      });
    } else {
      await logoutUser();
    }
  }

  await checkAccessTokenExpiry();
  if(accessToken == '' || accessTokenStatus == 'refreshing' || accessTokenStatus == 'expired') {
    // setTimeout(() => {
    //   console.log("Delayed for 250 milli second in new presentation.");
    //   create();
    // }, "800");
    return create();
  } else if(accessTokenStatus == 'refreshed') { 
    create(); 
  }
  
}


/* Add content to the exisiting slide in a Presentation */
const addToExistingSlides = async (presentationId, info, isFirst = true, isImage = false, isImageInserted = false) => {
    let textmsg = '', layout = 'TITLE_AND_BODY';
    if(info?.selectionText) {
      textmsg += info.selectionText;
    }
    // If only first slide is present start insertion from the second slide
    // Leave first slide for title and sub title alone 
    let result = await getSlideInsertionPositions(presentationId);
    if(result.slideLength<=0) 
      // || (result.slideStartIndex === undefined && isFirst))
    {
      if(info?.mediaType === "image") layout = 'BLANK';
      result = await createNewSlide(presentationId, layout);
      isFirst = false;
      // result = await getSlideInsertionPositions(presentationId);
    }
    if(info?.mediaType === "image" && !isImage) {
      if(isFirst) {
        result = await createNewSlide(presentationId, 'BLANK');
        // result = await getSlideInsertionPositions(presentationId);
      } 
      await insertImageToLastSlide(presentationId, info, result.slideDetails);
      isImageInserted = true;
    }
    if(info?.selectionText) {
      if (result?.slideEndIndex > 690 || result?.slideEndIndex+(textmsg.length) > 690 || isImageInserted || isEmpty(result.lastTextBoxId)) {
        result = await createNewSlide(presentationId, 'TITLE_AND_BODY');
        // result = await getSlideInsertionPositions(presentationId);
        await insertTextToLastSlide(presentationId, result?.lastTextBoxId, info, result?.slideStartIndex || 0);
      } else {
        await insertTextToLastSlide(presentationId, result?.lastTextBoxId, info, result?.slideStartIndex || 0);
      }
    }   
}

function getFormattedNames() {
  // Get current date and time
  const now = new Date();

  // Format date components
  const day = String(now.getDate()).padStart(2, '0');
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const year = now.getFullYear();

  // Format time components
  const hours = String(now.getHours() % 12 || 12).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  const seconds = String(now.getSeconds()).padStart(2, '0');
  const meridiem = now.getHours() < 12 ? 'AM' : 'PM';

  // Construct the formatted string
  const formattedSlideName = `SaveToGSlides ${day}-${month}-${year} ${hours}:${minutes}:${seconds} ${meridiem}`;
  const formattedSingleSlideName = `newSlide_${minutes}_${seconds}_${meridiem}`;
  return {
    formattedSlideName,
    formattedSingleSlideName,
  };
}

const addTitleAndSubtitleInFirstSlide = async (data, info) =>{
  const presentationId = data.presentationId;
  const postSlideApiUrl = `https://slides.googleapis.com/v1/presentations/${presentationId}:batchUpdate`;
  let textmsg = '';
  const requests = [];
  const titleSlide = data['slides'][0];
  const titleId = titleSlide['pageElements'][0]['objectId'];
  const subtitleId = titleSlide['pageElements'][1]['objectId'];
  requests.push(  {
    "insertText": {
      "objectId": titleId,
      "text": `Save to GSlides\n`
    }
  },
  {
    "insertText":{
      "objectId":subtitleId,
      "text": `Edit your google slides from anywhere without even opening`
    }
  }
);
  const create = async(info) => {
    if(accessTokenStatus == 'refreshed') {
      await fetch(postSlideApiUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({"requests":requests})
      })
      .then(response => response.json())
      .then(async data => {
        if (!data?.error) await addToExistingSlides(data.presentationId, info);
      })
      .catch(error => {
        console.error('Error:', error);
      });
    } else {
      await logoutUser();
    }
  }

  await checkAccessTokenExpiry();
  if(accessToken == '' || accessTokenStatus == 'refreshing' || accessTokenStatus == 'expired') {
    // setTimeout(() => {
    //   console.log("Delayed for 250 milli second in new presentation.");
    //   create(info);
    // }, "800");
    return create(info);
  } else if(accessTokenStatus == 'refreshed') { 
    create(info); 
  }
};

/* Creates New Presentation */
const createNewPresentation = async (info) => {
  const createSlideUrl = 'https://slides.googleapis.com/v1/presentations';

  let newSlideData = {};
  const create = async() => {
    if(accessTokenStatus == 'refreshed') {
      const { formattedSlideName } = getFormattedNames();
      await fetch(createSlideUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          "title": formattedSlideName
        })
      })
        .then(response => response.json())
        .then(async data => {
          await addTitleAndSubtitleInFirstSlide(data, info);
          await handleClick(info);
        })
        .catch(error => {
          console.error('Error:', error);
        });
  
      return newSlideData;
    } else {
      await logoutUser();
     }
  }

  await checkAccessTokenExpiry();
  if(accessToken == '' || accessTokenStatus == 'refreshing' || accessTokenStatus == 'expired') {
    // setTimeout(() => {
    //   console.log("Delayed for 250 milli second in new presentation.");
    //   create();
    // }, "800");
    return create();
  } else if(accessTokenStatus == 'refreshed') { 
    create(); 
  }
}


/* All Listeners*/

/* Menu Listeners */
chrome.contextMenus.onClicked.addListener(async function (info, tab) {
  if (info.menuItemId === "newPresentation") {
    await createNewPresentation(info);
  } else if (info.menuItemId === "ShowMore") {
    const presentationList = await getAllGoogleSlideList();
    chrome.contextMenus.remove('ShowMore', async () => {
      await createChildContextMenu(presentationList, pageCount, itemsPerPageCount, true);
    });
  } else if (info.parentMenuItemId === "ExistingSlide") {
    const menuItemId = info.menuItemId;
    await addToExistingSlides(menuItemId, info);
    setTimeout(() => {
      handleClick(info);
    }, "250");
   
  } else if (info.menuItemId === "mainMenu") {
    if (user_signed_in) {
      console.log("User is already signed in.");
    } else {
      const AUTH_URL = constructAuthUrl();
      loginUser(AUTH_URL);
    }
  } else if (info.menuItemId === "Recent") {
    chrome.storage.local.get(['Recent'], function(result) {
      const recentValue = result['Recent'];
      addToExistingSlides(recentValue, info);
    });
  }
});

/* Script Listeners */
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.message === 'login') {
    if (user_signed_in) {
      console.log("User is already signed in.");
      sendResponse('success');
    } else {
      let AUTH_URL = constructAuthUrl();
      try {
        loginUser(AUTH_URL);
        sendResponse('success');
      } catch (error) {
        console.error('Error during login:', error);
        sendResponse('error');
      } finally {
        sendResponse('success');
      }  
    }
  } else if (request.message === 'logout') {
    user_signed_in = false;
    logoutUser();
    sendResponse('success');
  }
});

chrome.runtime.onInstalled.addListener(function () {
  createMainMenu();
});
