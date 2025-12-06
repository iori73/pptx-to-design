// UI script for the Figma plugin

let selectedFile: File | null = null;

// DOM Elements
const dropZone = document.getElementById('dropZone') as HTMLDivElement;
const fileInput = document.getElementById('fileInput') as HTMLInputElement;
const fileInfo = document.getElementById('fileInfo') as HTMLDivElement;
const fileName = document.getElementById('fileName') as HTMLDivElement;
const fileSize = document.getElementById('fileSize') as HTMLDivElement;
const importBtn = document.getElementById('importBtn') as HTMLButtonElement;
const progress = document.getElementById('progress') as HTMLDivElement;
const progressFill = document.getElementById('progressFill') as HTMLDivElement;
const progressText = document.getElementById('progressText') as HTMLDivElement;
const status = document.getElementById('status') as HTMLDivElement;
const separateFrames = document.getElementById('separateFrames') as HTMLInputElement;
const preserveNames = document.getElementById('preserveNames') as HTMLInputElement;

// Drop zone events
dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  
  const files = e.dataTransfer?.files;
  if (files && files.length > 0) {
    handleFileSelect(files[0]);
  }
});

// File input change
fileInput.addEventListener('change', () => {
  const files = fileInput.files;
  if (files && files.length > 0) {
    handleFileSelect(files[0]);
  }
});

// Handle file selection
function handleFileSelect(file: File): void {
  if (!file.name.toLowerCase().endsWith('.pptx')) {
    showStatus('Please select a .pptx file', 'error');
    return;
  }
  
  selectedFile = file;
  
  // Show file info
  fileName.textContent = file.name;
  fileSize.textContent = formatFileSize(file.size);
  fileInfo.classList.add('visible');
  importBtn.classList.add('visible');
  
  // Hide any previous status
  status.classList.remove('visible');
}

// Format file size
function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

// Import button click
importBtn.addEventListener('click', async () => {
  if (!selectedFile) return;
  
  importBtn.disabled = true;
  progress.classList.add('visible');
  status.classList.remove('visible');
  updateProgress(0, 'Reading file...');
  
  try {
    const arrayBuffer = await selectedFile.arrayBuffer();
    
    updateProgress(10, 'Sending to plugin...');
    
    // Send to plugin
    parent.postMessage({
      pluginMessage: {
        type: 'import-pptx',
        data: Array.from(new Uint8Array(arrayBuffer)),
        options: {
          separateFrames: separateFrames.checked,
          preserveNames: preserveNames.checked,
        },
      },
    }, '*');
  } catch (error) {
    showStatus(`Error reading file: ${error}`, 'error');
    importBtn.disabled = false;
    progress.classList.remove('visible');
  }
});

// Update progress
function updateProgress(percent: number, text: string): void {
  progressFill.style.width = `${percent}%`;
  progressText.textContent = text;
}

// Show status message
function showStatus(message: string, type: 'success' | 'error'): void {
  status.textContent = message;
  status.className = `status visible ${type}`;
}

// Listen for messages from the plugin
window.onmessage = (event) => {
  const message = event.data.pluginMessage;
  if (!message) return;
  
  switch (message.type) {
    case 'progress':
      updateProgress(message.percent, message.text);
      break;
    case 'success':
      progress.classList.remove('visible');
      showStatus(message.message, 'success');
      importBtn.disabled = false;
      break;
    case 'error':
      progress.classList.remove('visible');
      showStatus(message.message, 'error');
      importBtn.disabled = false;
      break;
  }
};

