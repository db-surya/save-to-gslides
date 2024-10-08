document.addEventListener('DOMContentLoaded', function () {
  document.getElementById('loginButton').addEventListener('click', function () {
    chrome.runtime.sendMessage({ message: 'login' }, async function (response) {
      if (response === 'success') { 
        window.close();
      } else {
        window.close();
      }
    });
  });

  

});