const GAS_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbyNSu_HlAemXEdVjaxeNu-m15Uln5qBzv4-ZfnoyoIWKCbCuAfuLN1AnVX9s9zgxuuj/exec';

window.google = window.google || {};
window.google.script = window.google.script || {};

window.google.script.run = new Proxy({}, {
  get: function(target, prop) {
    if (prop === 'withSuccessHandler') {
      return function(handler) { return createRunner(handler, null); };
    }
    if (prop === 'withFailureHandler') {
      return function(handler) { return createRunner(null, handler); };
    }
    return function(...args) { executeApiCall(prop, args, null, null); };
  }
});

function createRunner(successHandler, failureHandler) {
  const runner = new Proxy({}, {
    get: function(target, prop) {
      if (prop === 'withSuccessHandler') {
        return function(handler) { successHandler = handler; return runner; };
      }
      if (prop === 'withFailureHandler') {
        return function(handler) { failureHandler = handler; return runner; };
      }
      return function(...args) { executeApiCall(prop, args, successHandler, failureHandler); };
    }
  });
  return runner;
}

function executeApiCall(action, args, onSuccess, onFailure) {
  fetch(GAS_WEB_APP_URL, {
    method: 'POST',
    body: JSON.stringify({ action: action, args: args }),
    headers: { 'Content-Type': 'text/plain;charset=utf-8' }
  })
  .then(res => res.json())
  .then(result => {
    if (result.error) {
      if (onFailure) onFailure(new Error(result.error));
      else console.error('API Error:', result.error);
    } else {
      if (onSuccess) onSuccess(result.data);
    }
  })
  .catch(err => {
    if (onFailure) onFailure(err);
    else console.error('Network Error:', err);
  });
}
