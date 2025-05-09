
// when a file is chosen, submit the form
document.querySelector("input[type=file]").addEventListener("change", function(){
	document.querySelector("form").submit();

    // make the box green
    document.querySelector(".file").style.backgroundColor = "#12a753";
    
    // change the label text
    document.querySelector("label").innerHTML = "File sucessfully converted and downloaded! Click to choose another file.";
});


// make the whole div clickable
document.querySelector(".file").addEventListener("click", function(){
    document.querySelector("input[type=file]").click();
});
 