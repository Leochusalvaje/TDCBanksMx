library("siebanxicor")
setToken("be72c8c74afd696286702192c2671f13e82737a362e413b5ad4ddaf661c90467")

idSeries <- c("SF235713","SF46410")
series <- getSeriesData(idSeries, '2020-01-01','2023-01-31')
head(series)

