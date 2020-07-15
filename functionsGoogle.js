//
//   Developer: J.A. Runnells
//     Updated: 2020-07-15 14:20
//   Last Push: 2020-07-15 14:20
//      Branch: master
//     License: MIT
//

/*
    Google Sheets:
        [1] https://www.makeuseof.com/tag/create-custom-functions-google-sheets/
        [2] https://developers.google.com/apps-script/guides/sheets/functions

        x̄  p̂  α
*/

/* ========================== UPDATED VERSIONS ========================== */

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
 * Calculates the margin of error [E or ME] for the mean.
 * @constructor
 * @param {number} t_alpha2 - t_α/2 [t sub alpha divided by 2]
 * @param {number} s - sample standard deviation
 * @param {number} n - sample size
 * @return The value of t_α divided by two, then multiplied by sample standard dev divided by the square root of sample size.
 * @customfunction
 */
function MARGIN_ERROR_M(t_alpha2,s,n) {
    return t_alpha2*(s/Math.sqrt(n));
}

/**
 * Calculates the margin of error [E or ME] for a proportion.
 * @constructor
 * @param {number} z_alpha2 - z_α/2 [z sub alpha divided by 2]
 * @param {number} phat - p̂ (population proportion)
 * @param {number} n - sample size
 * @return The value of z_α divided by two, then multiplied by σ_p̂
 * @customfunction
 */
function MARGIN_ERROR_P(z_alpha2,phat,n) {
  //return z_alpha2*Math.sqrt((phat*(1-phat))/n);
  return z_alpha2*STDEV_PHAT(phat,n);
}

/**
 * Calculates the specified upper or lower bound value.
 * @constructor
 * @param {number} z_alpha2 - z_α/2 [z sub alpha divided by 2]
 * @param {number} phat - p̂ (population proportion)
 * @param {number} n - sample size
 * @param {boolean} lower - false for lower bound (default), true for upper bound.
 * @return The confidence interval
 * @customfunction
 */
function CONFIDENCE_INTERVAL(z_alpha2,phat,n,lower=false) {
    switch (lower) {
        case false:
            return phat - MARGINERROR(z_alpha2,phat,n);
        case true:
            return phat + MARGINERROR(z_alpha2,phat,n);
    }
}