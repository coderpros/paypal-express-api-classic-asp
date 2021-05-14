
/**
 * Events
 */
$("#AddButton").click(function() {
    $("#store-container tbody > tr:first").clone().appendTo("#store-container tbody").show("slow");
});

$("#RemoveButton").click(function() {
    $("#store-container tbody > tr:last").remove();
});

$("#SubmitButton").click(function() {
    $("#store-container tbody > tr:first").remove();
});