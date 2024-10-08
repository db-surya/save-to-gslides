document.addEventListener('DOMContentLoaded', function () {
  document.getElementById('logoutButton').addEventListener('click', function () {
    chrome.runtime.sendMessage({ message: 'logout' }, function (response) {
      if (response === 'success') window.close();
      else window.close();
    });
  });

});