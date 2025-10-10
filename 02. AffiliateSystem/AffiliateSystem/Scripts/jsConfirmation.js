function ConfirmedDelete(e) {
    //Confirmation delete
    var msg = confirm('Are you sure want to delete this data ?');
    if (msg == false) {
        e.processOnServer = false;
        return;
    }
}

function ConfirmedSubMenu(e) {
    //Confirmation delete
    var msg = confirm('Are you sure want to close this menu ?');
    if (msg == false) {
        e.processOnServer = false;
        return;
    }
}