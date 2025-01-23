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

  const isAtBottom = logContainer.scrollHeight - logContainer.scrollTop - logContainer.clientHeight <= 5;

  logContainer.appendChild(logEntry);

  if (isAtBottom) {
    logContainer.scrollTop = logContainer.scrollHeight;
  }
  
});

document.getElementById('open-instructions').addEventListener('click', () => {
  ipcRenderer.send('open-instructions');
});

document.getElementById('open-second').addEventListener('click', () => {
  ipcRenderer.send('open-second-window');
});