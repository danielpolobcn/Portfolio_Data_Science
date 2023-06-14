# INSERT SEARCH PARAMETERS BELOW:
# Radius in Km, date-time in format: YYYY-MM-DDTHH:mm:ssZ).
lat = "42.361145"
lon = "-71.057083"
radius = "10"
start_time = "2022-01-01T00:00:00Z"
end_time = "2022-01-01T02:00:00Z"

# INSERT BEARER TOKEN BELOW:
bearer_token = "YOU NEED A TWITTER API ACADEMIC ACCOUNT TO PERFORM ARCHIVE SEARCHES"

# QUERY TO PASS TO TWITTER API:
query = "has:geo point_radius:["+lon+" "+lat+" "+radius+"km]"
query_params = {"query": query, "max_results": 500, "tweet.fields":"id,author_id,created_at", "expansions":"geo.place_id", "place.fields":"id,full_name,geo", "start_time":start_time, "end_time":end_time}


# SCRIPT START
import requests
import time
import pandas as pd

# build request for Twitter API with authentication.
search_url = "https://api.twitter.com/2/tweets/search/all"

def bearer_oauth(r):
    r.headers["Authorization"] = "Bearer " + bearer_token
    r.headers["User-Agent"] = "TwitterAPIv2_ArchiveSearchPython"
    return r

def connect_to_endpoint(url, params):
    response = requests.get(url, auth=bearer_oauth, params=params)
    print(response.status_code)
    if response.status_code != 200:
        raise Exception(response.status_code, response.text)
    return response.json()

# main function: send request and build dataset, output as CSV.
def main():
    response = connect_to_endpoint(search_url, query_params)
    
    # calculate centroid of bounding box:
    def find_centroid(lon1, lat1, lon2, lat2):
          lon = (lon1 + lon2) / 2
          lat = (lat1 + lat2) / 2
          return lon, lat
    
    # retrieve full name and bounding box for "place_id".
    def get_place_data (place):
        for n in response["includes"]["places"]:
            if n["id"] == place:
                return n["full_name"], find_centroid(*n["geo"]["bbox"]) 
    
    # extract data from object response and merge tweets.
    def merge_tweets(tweets, data):
        for x in data:
            tweet = {}
            tweet["id"] = x["id"]
            tweet["author_id"] = x["author_id"]
            tweet["created_at"] = x["created_at"]
            try:
                tweet["place_id"] = x["geo"]["place_id"]
            except:
                tweet["place_id"] = ""
            if tweet["place_id"] != "" and "includes" in response and "places" in response["includes"]:
                tweet["place_name"] = get_place_data(x["geo"]["place_id"])[0]
                tweet["longitude"] = get_place_data(x["geo"]["place_id"])[1][0]
                tweet["latitude"] = get_place_data(x["geo"]["place_id"])[1][1]
            else:
                tweet["place_name"] = ""
                tweet["longitude"] = ""
                tweet["latitude"] = ""
            try:
                tweet["text"] = x["text"]
            except:
                tweet["text"] = ""
            tweets.append(tweet)
        return tweets
    
    tweets = []
    data = []
    if("data" in response):
        data = response["data"]
        tweets = merge_tweets(tweets, data)
            
    # iteration until no next_token is found (end of pagination).
    if "meta" in response and "next_token" in response["meta"]:
        next_token = response["meta"]["next_token"]
        while True:
            time.sleep(3) #api rate limit 300 per 15 min
            query_params["next_token"] = next_token
            response = connect_to_endpoint(search_url, query_params)
            if("data" in response):
                data = response["data"]
                tweets = merge_tweets(tweets, data)
                if("meta" in response and "next_token" in response["meta"]):
                    next_token = response["meta"]["next_token"]
                else:
                    break
    
    df = pd.DataFrame(tweets)
    df.to_csv("twitter_search.csv")
        
if __name__ == "__main__":
    main()
#SCRIPT END
