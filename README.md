# YearlyStockMarketAnalysis
Using Excel and VBA to analyze stock market trends year over year 

## StockSummarizer
This is the main module that will summarize all the ticker data for a given year. It utilizes two class modules to assist with manipulating the data. In this module, each ticker
symbol will be summarized by it's percent change, price change, and total volume. The summaries will also be formatted as green if there was a positive change or red
if the change was negative.

## StockDaily
Object that holds the data for a single day of provided data for a ticker. There are also two functions in the class that will calculate the price change and percent change
based on another daily that is passed in as a parameter.

## StockSummary
Class that holds the data to summarize a ticker over a year. A collection of these are made to store and then display the summaries on each worksheet
