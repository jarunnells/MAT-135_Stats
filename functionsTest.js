/*
    Google Sheets:
        [1] https://www.makeuseof.com/tag/create-custom-functions-google-sheets/
        [2] https://developers.google.com/apps-script/guides/sheets/functions

        x̄  p̂  α
*/

/*===========================================================================*/

/**
 * Calculates standard deviation of the sampling distribution of p̂ (population proportion) from given X and n.
 * @constructor
 * @param {number} p - population proportion.
 * @param {number} n - sample size
 * @return The square root of the product of population proportion and its complement divided by the sample size.
 * @customfunction
 */
function STDEV_PHAT(p,n) {
    return Math.sqrt((p*(1-p))/n);
}

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
 * @param {number} z_alpha2 - z_α/2 [z sub alpha divided by 2]
 * @param {number} phat - p̂ (population proportion)
 * @param {number} n - sample size
 * @return The value of z_α divided by two, then multiplied by
 * @customfunction
 */
function MARGIN_ERROR(z_alpha2,phat,n) {
    return z_alpha2*Math.sqrt((phat*(1-phat))/n);
}

/**
 * Calculates the specified upper or lower bound value.
 * @constructor
 * @param {number} z_alpha2 - z_α/2 [z sub alpha divided by 2]
 * @param {number} phat - p̂ (population proportion)
 * @param {number} n - sample size
 * @param {boolean} lower - true for lower bound, false for upper bound.
 * @return The
 * @customfunction
 */
function CONFIDENCE_INTERVAL(z_alpha2,phat,n,lower) {
    switch (lower) {
        case true:
            return phat - MARGINERROR(z_alpha2,phat,n);
        case false:
            return phat + MARGINERROR(z_alpha2,phat,n);
    }
}
/**/
