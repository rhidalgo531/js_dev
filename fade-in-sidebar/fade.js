document.addEventListener("DOMContentLoaded", main);

function main() {
  $(document).scroll(function() {
   $("#sideBar").hide().fadeIn(200);
    $('#sideBar').css("top", $(document).scrollTop() + 100);
  });
}
