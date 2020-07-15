/*

    Google Sheets:
        [1] https://www.makeuseof.com/tag/create-custom-functions-google-sheets/
        [2] https://developers.google.com/apps-script/guides/sheets/functions

*/
// x̄  p̂  α 


function phat(x,n) {

    var phat = x/n;

    return phat;
}

function marginErr(z_alpha,phat,n) {

    var marginErr = (z_alpha/2)*Math.sqrt((phat*(1-phat))/n);

    return marginErr;
}

/*===========================================================================*/

/**
 * Calculates p̂ (population proportion) from given X and n.
 * @constructor
 * @param {number} x - individuals in the sample with a specified characteristic.
 * @param {number} n - sample size
 * @return Number of individuals divided by sample size.
 * @customfunction 
 */
function PHAT(x,n) {    

    return x/n;
}

/**
 * Calculates the margin of error [E or ME] for a proportion.
 * @constructor
 * @param {number} z_α - [z sub alpha] 
 * @param {number} p̂ - [p hat] 
 * @param {number} n - sample size
 * @return The value of z_α divided by two, then multiplied by the square root of the product of p hat and the complement of p hat divided by the sample size.
 * @customfunction
 */
function MARGINERROR(z_alpha2,phat,n) {

    return z_alpha2*Math.sqrt((phat*(1-phat))/n);
}
