# baseball_scouting

This is a baseball scouting/projection tool. It can be used to forecast prospects future success, and/or forecast current MLB players future years. I am building this for my own keeper fantasy league, which does NOT care about defense. Thus this project also does not care about defense. We do scrape defensive stats if you are interested. I may include this in the future, but at the moment I am not

Using publicly accessable API's, you can scrap the internet for baseball stats of all levels. I'm only interested in AA, AAA and MLB so that is all is being scraped at the moment. However you can scrape all the way down to Rookie is you so wish.
	To scrape the data, open the raw data folder and run getData.py. Under the assumption you've got the required packages installed, it should run and create (if they are not already there) a hitter and pitcher subfolder in the raw data folder. Then it will input all data from the requested level(s) for each unique play in the time frame chosen.

PROSPECT PROJECTION
Begin by running cleanProspect.py, as this will clean the raw data into the data needed for prospect projection. This is for players whom have never played an MLB game, or very few. As prospects do not have the advanced stat trackers like the MLB does, its unnecessary complexity to have MLB stats that MiLB does not have. I am aware of prospect savant, but its fan run and NOT public (as of Nov 2026). I may reach out to the admin to ask for permission to scrap their website, but I have not as of yet.

	FUTURE PLAN
		Build a GLM/statistical model that finds the most influencial/statistically significant stats.
		Build a time series ml model that uses the statistically significant stats to project an inputted MiLB player.
			You would be prompted to enter a Players name and orginization, so that we can scrap their data and compare them to MLB players career.
			
		Output
			Contain their entire projected MLB career (best estimate)
			Contain a ceiling career projection (maybe HOF worthy, maybe not)
			Contain a worst case career projection
			Contain a list of top 5 most similar players

MLB PROJECTION
Begin by running cleanMLB.py, as this will eliminate a players begining MiLB before making their MLB debut. If we are projecting a fringe player (up and down), it keeps their MiLB stats after their MLB debut. This allows us to use advanced statistics that are only available to MLb players (at them moment).

	FUTURE PLAN
		Build a GLM/statistical model that finds the most influencial/statistically significant stats.
		Build a time series ml model that uses the statistically significant stats to project an inputted MLB player.
			You would be prompted to enter a Players name and orginization, so we know who to remove from the training dataset. To then test them against the rest of the MLB.
		
		OutPut
			Contain the rest of their projected career (best estimate)
			Contain a best case and worse case for the remainder of their career
			* Not sure how im going to implement it yet, but the output will include a salary guide for this player
				Essentially answering, What AAV and length of contract are they worth each year?
