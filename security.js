/**
 * Website Protection & Security Script
 * Prevents text selection, copying, image dragging, and screenshot utilities (like Snipping Tool).
 */
document.addEventListener('DOMContentLoaded', function() {
  // Create and append the screenshot overlay div dynamically
  const overlay = document.createElement('div');
  overlay.id = 'screenshot-overlay';
  overlay.style.position = 'fixed';
  overlay.style.top = '0';
  overlay.style.left = '0';
  overlay.style.width = '100%';
  overlay.style.height = '100%';
  overlay.style.background = 'black';
  overlay.style.zIndex = '999999';
  overlay.style.display = 'none';
  overlay.style.pointerEvents = 'none';
  document.body.appendChild(overlay);

  // When window loses focus (like Snipping Tool or screenshot software active), show black screen
  window.addEventListener('blur', function() {
    overlay.style.display = 'block';
  });

  // When window gains focus back, hide black screen
  window.addEventListener('focus', function() {
    overlay.style.display = 'none';
  });

  // Print screen key detection
  window.addEventListener('keyup', function(e) {
    if (e.key === 'PrintScreen' || e.keyCode === 44) {
      overlay.style.display = 'block';
      setTimeout(function() {
        overlay.style.display = 'none';
      }, 1000);
    }
  });

  // Extra copy protection: prevent browser's copy/cut events
  document.addEventListener('copy', function(e) {
    e.preventDefault();
  });
  document.addEventListener('cut', function(e) {
    e.preventDefault();
  });
});
