/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/

// "use strict";

(function () {

  // Function for infinite scroll to load images progressively
  function infiniteScroll () {
    this.initialize = function() {
      this.setupEvents();
    };
  
    this.setupEvents = function() {
      $('#content-main').on('scroll',
        this.handleScroll.bind(this)
      );
    };
  
    var throttleTriggered = false;
    this.handleScroll = function() {
      var scrollTop = $('#content-main').scrollTop();
      var windowHeight = $(window).height();
      var height = $('#images-list').height() - windowHeight;
      var scrollPrecentage = ((scrollTop / height));
      if ( !throttleTriggered ) {
        throttleTriggered = true;
        setTimeout(function(){
          throttleTriggered = false;
          if(scrollPrecentage > 0.9 ) {
            showImages(currentQuerry);
          }
        }, 700);
      }
    };
    this.initialize();
  }

  // Variables used for http query management.
  var newSearch = false;
  var basicQuerry = 'https://pixabay.com/api/?key=7382177-e4b9357cd416cbd1674bfe1d9';
  var currentQuerry = basicQuerry;
  var searchPage = 1;
  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      var searchBar = $('#search-bar');
      var searchForm = $('#search-form');
      searchBar.on("input", function(){
        getSearchResults($(this).val());
      });
      searchForm.on('submit', function(event){
        event.preventDefault();
        getSearchResults($(searchBar).val());
      });
      // Initilize Images
      showImages(basicQuerry);
      // Initialize scroll
      infiniteScroll();
    });
  };


  // Get search results
  {
    var debounceTimer;
  }
    function getSearchResults(searchTags) {
      words = searchTags.trim().split(" ");
      var querryString = basicQuerry + '&q=' + words.join('+');
      currentQuerry = querryString;
      newSearch = true;
      clearTimeout(debounceTimer);
  
      debounceTimer = setTimeout(function(){
        showImages(querryString);
      }, 1000);
    }
// Get Images from the API
  function showImages (url) {
    // var adress =  + url;
  $.ajax({
      url: url + '&page=' + searchPage,
      type: 'GET',
      crossDomain: true,
      success: function (result){
        var hits = result.hits;
        var imagesList = $('#images-list');
        if(newSearch) {
          imagesList = $('<ul id="images-list"></ul>');
          $('#images-list').remove();
          newSearch = false;
          searchPage = 0;
        }
        searchPage++;
        for( var i=0; i <= hits.length; i++) {
          var item = hits[i];
          if(item){
            var listItem = $('<li class="ms-ListItem image-thumbnail"><img class="mini-image" src=' + item.webformatURL + ' alt=' + item.id + ' crossOrigin="Anonymous"/></li>');
            listItem.click(function(event) {
              insertImage(event.currentTarget.firstChild.src);
            });
            imagesList.append(listItem);       
          }
        }
        var imageContainer = $('#images');
        imageContainer.append(imagesList);
      },
      error: function (xhr, status, error) {
        console.log("Something went wrong", xhr, status, error);
      }
    });
  }

//Inserting image as base64 string
   function insertImage (url) {
   
    toBase64(url).then(function (element){ 
      Office.context.document.setSelectedDataAsync(element, {
         coercionType: Office.CoercionType.Image,
         },
         function (asyncResult) {
           if (asyncResult.status === Office.AsyncResultStatus.Failed){
             console.log(asyncResult.error.message);
           }
         }  
       );
    });
   }
 
// Change picture to base64 string.

  function toBase64(url){
      return new Promise(function (resolve, reject){
        var canvas = document.createElement('canvas');
        var context = canvas.getContext('2d');
        var image = new Image();
        image.crossOrigin = 'anonymous';
        image.onload = function () {
          canvas.height = image.naturalHeight;
          canvas.width = image.naturalWidth;
          context.drawImage(image, 0, 0);
          var data = canvas.toDataURL("image/png").replace(/^data:image\/(png|jpg);base64,/, "");
          if (data !== 'data:,'){
            resolve(data);
          } else {
            console.log("Need handle This!");
          }
        }
        image.src = url;
    });
  }
})();
