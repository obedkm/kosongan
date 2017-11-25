var pageCounter = 1;
var btn = document.getElementById('btn');
var animalDiv = document.getElementById('animal-info');

btn.addEventListener("click",function(){
	var ourRequest = new XMLHttpRequest();
	ourRequest.open('GET','https://learnwebcode.github.io/json-example/animals-'+pageCounter+'.json');
	ourRequest.onload = function() {
		// console.log(ourRequest.responseText);
		var ourData = JSON.parse(ourRequest.responseText);
		renderHTML(ourData);
		// console.log(ourData[0]);
	};
	ourRequest.send();
	pageCounter++;
	if (pageCounter > 3) {
		btn.classList.add("hide-me");
	}
})

function renderHTML(data) {
	var htmlString = "";
	for (var i = 0;i <data.length; i++) {
		htmlString +="<p> " + data[i].name + " is a " + data[i].species + ".</p>"; 
	}
	animalDiv.insertAdjacentHTML('beforeend',htmlString);


}
