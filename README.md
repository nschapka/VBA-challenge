# VBA-challenge
Module 2 challenge assignment

# Overview

Steve's parents are looking to invest some money into individual stocks.  As such, they asked Steve to recommend a few in which to invest.  Steve knows that I am a brain genius and code wizard, so he asked me to help with the analysis, to which I graciously agreed.

### Analysis

For a selection of 12 candidate stocks, we will look at the total trading volume and yearly return for 2017 and 2018.  To do so, we use a VBA script to iterate across the dataset, adding up daily trading volumes and grabbing start and end prices as we go.

When we have read all the necessary data from the spreadsheet, we present it in a concise table on a new worksheet.

# Results

### 2017

2017 was a good year for most of the candidate stocks.

![vba_challenge_2017](https://user-images.githubusercontent.com/38693762/148699222-a4cca8b8-89a6-4576-8871-43fdceccc4d1.PNG)

In particular, it would have been a very good idea to have invested in DQ or SEDG.

### 2018

Oh no! Looks like someone accidentally turned the big lever on the NYSE from "line go up" to "line go down"!

![vba_challenge_2018](https://user-images.githubusercontent.com/38693762/148699242-7d4e2ec8-6955-46c8-8452-cb19ff45017f.PNG)

The only two stocks with positive returns in both years are RUN and ENPH.  Within the limited scope of this analysis, RUN and ENPH seem to be the most resilient to market fluctuations, and thus the best out of the 12 candidates for Steve's parents to invest in.

### My advice

If Steve's parents are asking their son for stock advice, it would stand to reason that they are not rich enough to have the sort of insider connections necessary to make short-term individual stock investments worthwhile.  Additionally, if they are close to, or already retired, stocks are likely not a sound option anyway.  From my perspective as a 27 year old, when the market dips, my 401(k) is simply purchasing shares at a discount.  The total value of my portfolio will not matter for another 20 years at least.  Being a good friend, I do not want Steve's parents to end up like Jerry in S1E5 of Seinfeld, so I cannot in good conscience recommend any of the 12 stocks.

Bonds, however, are not subject to the extreme yearly fluctuations seen in the stock market, and are a much safer option for Steve's parents.

If Steve's parents have their hearts set on investing in the market, I would recommend holding on to an ETF that tracks the DJIA for a few years, or making friends with someone whose first name is "senator".

# Summary

### Refactoring?

Refactoring our original code allowed it to run much faster, but did not really expand the scope of what it could do.  For a one time limited analysis like this, it would be a more efficient use of time to simply use the old, slowly running code to perform the analysis, rather than take the effort to refactor the working code to execute more quickly.

The code, as it is now, could be pretty easily retooled a little more to be able to handle a much larger dataset, including hundreds or even thousands of tickers.  In this use case, manually running the old macro for each of a thousand tickers would take forever, and refactoring the code makes much more sense.

Additionally, if Steve and I were to open up a boutique stock trading and advice firm, we may be running such analyses thousands of times each day.  In this use case, refactoring is very important, to allow us to act as quickly as possible.  We may even want to abandon VBA altogether for a more efficient programming language.

