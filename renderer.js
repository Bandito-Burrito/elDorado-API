const { ipcRenderer } = require('electron')

document.getElementById('selectFile').addEventListener('click', async () => {
  const filePath = await ipcRenderer.invoke('open-file-dialog')
  if (filePath) {
    console.log('Selected file path:', filePath)
    document.getElementById('logger').textContent = `Selected file: ${filePath}`

    ipcRenderer.send('selected-file', filePath)
    
  } else {
    console.log('No file selected')
    document.getElementById('logger').textContent = 'No file selected'
  }
})

ipcRenderer.on('log', (event, message) => {
  const logContainer = document.getElementById('logger');
  const logEntry = document.createElement('div');

  logEntry.innerHTML = message.replace(/\n/g, '<br>');

  logContainer.appendChild(logEntry);

  // Use a timeout to let the DOM update before checking the scroll position
  setTimeout(() => {
    // Check if the user is at the bottom (or near the bottom)
    const isAtBottom = logContainer.scrollHeight - logContainer.scrollTop - logContainer.clientHeight <= 5;

    if (isAtBottom) {
      // Auto-scroll to the bottom
      logContainer.scrollTop = logContainer.scrollHeight;
    }
  }, 0);

});

document.getElementById('open-instructions').addEventListener('click', () => {
  ipcRenderer.send('open-instructions');
});
