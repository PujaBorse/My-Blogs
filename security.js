/**
 * Website Protection & Security Script
 * Prevents text selection, copying, image dragging.
 */
document.addEventListener('DOMContentLoaded', function() {
  // Extra copy protection: prevent browser's copy/cut events
  document.addEventListener('copy', function(e) {
    e.preventDefault();
  });
  document.addEventListener('cut', function(e) {
    e.preventDefault();
  });
});
