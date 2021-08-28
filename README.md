# VBA of Wall Street
## *Overview*
Originally Steve and his parents wanted a in-depth analysis of the DQ (Daqo) stock in 2018. With some initial research, they realized DQ might not be the smartest investment and wanted to expand the analysis to include all of the stocks in our data set in order to make the most informed investment decision. Lastly, they want to be able to use the code we're developing to analyze the entire stock market over the last few years, so we need to streamline the code to maximize efficiency for any future projects. 
### *Results*
At a glance, it's clear that the stocks in our dataset from 2017 (Figure A) ran significantly better than 2018 (Figure B). In 2017, 92% of the stocks had positive growth, 4 instances of which, they doubled at minimum. In 2018, however, only 2 stocks had positive growth: RUN and ENPH. So while RUN only grew 5.5% in 2017 (one of the lower-performing stocks that year), they had the largest growth (84%) in 2018, making it a strong option overall. The most consistent stock from these 2 years would have to be ENPH: it grew an impressive 130% in 2017 and 82% again in 2018. 
#### Figure A:
![All_Stocks_2017](https://user-images.githubusercontent.com/87578449/131223897-f0c06dfb-f55c-4241-bccf-c06ff821ad43.png)
#### Figure B:
![All_Stocks_2018](https://user-images.githubusercontent.com/87578449/131223905-e623c2c0-fc22-430a-b395-344db9208283.png)

As far as the execution time of the code, we were able to cut the time down from 0.984 seconds for 2017 to 0.246 seconds (Figure C). For 2018, similarly we were able to cut down the execution time from 0.828 seconds to 0.253 seconds (Figure D). While those split seconds may not seem like much, it runs about 4 times faster, which will make a huge difference once the dataset is expanded to include thousands of stocks over several more years. In order to increase the run time, we created 3 additional arrays (Figure E) instead of having another For Loop inside the original For Loop (Figure F).
#### Figure C:
![VBA_Challenge_2017](https://user-images.githubusercontent.com/87578449/131224071-648ee656-8936-405e-bd7f-da029c6015a1.png)
#### Figure D: 
![VBA_Challenge_2018](https://user-images.githubusercontent.com/87578449/131224073-5b863c53-9140-47ee-a1fa-63e7047523f1.png)
#### Figure E:
![Output_Arrays](https://user-images.githubusercontent.com/87578449/131224271-752d9f24-322f-441b-82b0-4b6920368fcd.png)
#### Figure F:
![For_Loops](https://user-images.githubusercontent.com/87578449/131224278-1d6cd97f-4961-422f-b931-39e669b7c8e5.png)

## *Summary*
