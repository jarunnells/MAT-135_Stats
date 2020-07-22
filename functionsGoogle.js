//
//   Developer: J.A. Runnells
//     Updated: 2020-07-22 00:45
//   Last Push: 2020-07-22 00:45
//      Branch: master
//     License: MIT
//

// Google Sheets: https://developers.google.com/apps-script/guides/sheets/functions
// x̄  p̂  α

// =============================================================================
// AVAILABLE FUNCTIONS (JS):
//   [01] STDEV_PHAT() <=> [x] TESTING  [x] DEV COMPLETE
//   [02] PHAT() <=> [x] TESTING  [x] DEV COMPLETE
//   [03] PHAT_I() <=> [x] TESTING  [x] DEV COMPLETE
//   [04] MARGINERROR_M() <=> [x] TESTING  [x] DEV COMPLETE
//   [05] MARGINERROR_P() <=> [x] TESTING  [x] DEV COMPLETE
//   [06] MARGINERROR_I() <=> [x] TESTING  [x] DEV COMPLETE
//   [07] CONFIDENCE_INTERVAL() <=> [x] TESTING  [x] DEV COMPLETE
//   [08] SAMPLE_MIN_P() <=> [x] TESTING  [x] DEV COMPLETE
//   [09] SAMPLE_MIN() <=> [x] TESTING  [x] DEV COMPLETE
//   [10] SAMPLE_MIN_MEAN() <=> [x] TESTING  [x] DEV COMPLETE
// =============================================================================

/**
 * Calculates standard deviation of the sampling distribution of p̂ (population proportion).
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
 * Calculates p̂ (sample proportion) given upper and lower bounds.
 * @constructor
 * @param {number} lower - lower bounds value
 * @param {number} upper - upper bounds value
 * @return The value of upper and lower bounds divided by two
 * @customfunction
 */
function PHAT_I(lower,upper) {
    return (lower+upper)/2;
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
 * Calculates the margin of error given upper and lower bounds.
 * @constructor
 * @param {number} lower - lower bounds value
 * @param {number} upper - upper bounds value
 * @return The difference between upper and lower bounds divided by twoSpreadsheetApp.getUi().alert
 * @customfunction
 */
function MARGIN_ERROR_I(lower,upper) {
    return Math.abs(lower-upper)/2;
}

/**
 * Calculates the specified upper or lower bound value.
 * @constructor
 * @param {number} z_alpha2 - z_α/2 [z sub alpha divided by 2]
 * @param {number} phat - p̂ (population proportion)
 * @param {number} n - sample size
 * @param {boolean} lower - lower - true (1) for lower bound [default], false (0) for upper bound.
 * @return The confidence interval ( sum of p̂ ± ME ).
 * @customfunction
 */
function CONFIDENCE_INTERVAL(z_alpha2,phat,n,lower=true) {
    switch (lower) {
      case true:
      case 1:
          return phat-MARGIN_ERROR_P(z_alpha2,phat,n);
      case false:
      case 0:
          return phat+MARGIN_ERROR_P(z_alpha2,phat,n);
      default:
          return "[ERROR]";
    }
}

/**
 * Estimates population proportion using prior estimate (p̂).
 * @constructor
 * @param {number} phat - p̂ (population proportion)
 * @param {number} z_alpha2 - z_α/2 [z sub alpha divided by 2]
 * @param {number} me - margin of error
 * @return The product of p̂, p̂', and the result of z_α/2 divided by the margin of error squared.
 * @customfunction
 */
 function SAMPLE_MIN_P(phat,z_alpha2,me) {
     return phat*(1-phat)*Math.pow((z_alpha2/me),2);
 }

 /**
  * Estimates population proportion without a prior estimate.
  * @constructor
  * @param {number} z_alpha2 - z_α/2 [z sub alpha divided by 2]
  * @param {number} me - margin of error
  * @return The product of 0.25 and z_α/2 divided by the margin of error squared.
  * @customfunction
  */
function SAMPLE_MIN(z_alpha2,me) {
    return 0.25*Math.pow((z_alpha2/me),2);
}

 /**
  * Estimates population mean.
  * @constructor
  * @param {number} z_alpha2 - z_α/2 [z sub alpha divided by 2]
  * @param {number} s - sample standard deviation
  * @param {number} me - margin of error
  * @return The product of z_α/2 and s divided by the margin of error squared.
  * @customfunction
  */
function SAMPLE_MIN_MEAN(z_alpha2,s,me) {
    return Math.pow((z_alpha2*s/me),2);
}
 