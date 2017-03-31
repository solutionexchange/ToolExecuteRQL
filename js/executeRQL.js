var SelBuffer = "";
var ScrollTop = 0;
function ExecuteRQL() {
	document.ioText.MarkedRQL.value = SelBuffer;
	document.ioText.ScrollTop.value = ScrollTop;
	document.ioText.submit();
}
function SaveSelection() {
	var objRange = document.selection.createRange();
	SelBuffer = objRange.text;
	ScrollTop = document.ioText.sRQLRequest.scrollTop;
}