$(function () {
  $("body").append("<div class='processingIcon'><img src='Images/processing.gif' style='height: 61px; width: 57px' /></div>");
  processingOff();
  $(".processingButton").click(function () {
    processingOn();
  });
})
function processingOff() {
  $("body").removeClass("processingCurtain");
  $(".processingIcon").hide();
}
function processingOn() {
  $("body").addClass("processingCurtain");
  $(".processingIcon").show();
  $(".processingButton").hide();
}