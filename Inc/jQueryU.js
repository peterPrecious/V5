/*
  source: //jquery-howto.blogspot.ca/2009/09/get-url-parameters-values-with-jquery.html

  ...to get object of URL parameters use:
  var allVars = $.getUrlVars();

  ...to get URL var by its name use:
  var byName = $.getUrlVar("name");
*/

$.extend({

  getUrlVars: function () {
    var vars = [], hash;
    var hashes = window.location.href.slice(window.location.href.indexOf('?') + 1).split('&');
    for (var i = 0; i < hashes.length; i++) {
      hash = hashes[i].split('=');
      vars.push(hash[0]);
      vars[hash[0]] = hash[1];
    }
    return vars;
  },

  getUrlVar: function (name) {
    return $.getUrlVars()[name];
  }

});