function validChar(e) {
    var unicode = e.charCode ? e.charCode : e.keyCode
    if (unicode != 8) { //if the key isn't the backspace key (which we should allow)
        if (unicode == 39) //if (')
            return false //disable key press
    }
}

function validNumeric(e) {
    var unicode = e.charCode ? e.charCode : e.keyCode
    if (unicode != 8) { //if the key isn't the backspace key (which we should allow)
        if ((unicode < 48 || unicode > 57) && (unicode > 36 || unicode < 41)) //if not a number
            return false //enable key press        
    }
}