import requests
import json
import sys
import os
import time;
from datetime import datetime
video_id = "";
if input("PRESS ENTER") != 0:
    URL = 'https://www.googleapis.com/youtube/v3/'
    PLAYLISTID = "UUv3DBWI1aZVYiq-WcARA83g";
    # ここにAPI KEYを入力
    API_KEY = ''
    try:
        assert API_KEY != '';
    except:
        print("ERROR: OPEN CODE AND SET YOUTUBE API KEY");
        input("PRESS ENTER TO EXIT");
        exit();

    params = {'key': API_KEY, 'part': 'snippet', 'playlistId': PLAYLISTID,}
    response = requests.get(URL + 'playlistItems', params=params);
    print(response)
    try:
        assert "error" not in response;
    except:
        input("AN ERROR OCCURRED")
        exit();
    video_id = response.json()["items"][0]["snippet"]["resourceId"]["videoId"];
    
    def getResult():
        try:
            f = open(video_id + ".csv", mode='a', encoding="utf-8");
            params = {
              'key': API_KEY,
              'part': 'statistics',
              'id': video_id,
            }
            response = requests.get(URL + 'videos', params=params);
            resourse = response.json();
            try:
                assert 'error' not in resourse;
            except:
                print(resourse['error']);
                f.close();
                sys.exit(input("An Error Occured."));
            dt_now = datetime.now()
            resourse = resourse["items"][0];
            viewCount = resourse["statistics"]["viewCount"];
            likeCount = resourse["statistics"]["likeCount"];
            commentCount = resourse["statistics"]["commentCount"];
            arr = [dt_now.strftime('%Y-%m-%d %H:%M:%S'), viewCount, likeCount, commentCount];
            txt = ["時刻", "再生数", "高評価数", "コメント数"];
            csv = ",".join(map(str, arr))+"\n";
            f.write(csv);
            for t, e in zip(txt, arr):
                print(t, e, end=" ");
            print();
            f.close();
        except:
            print("an error occured");
            f.close();
            sys.exit();
    while True:
        getResult();
        SEC = 60;
        while SEC:
            print("\r{0} ".format(SEC), end="");
            time.sleep(1);
            SEC-= 1;
        print("\r", end="");
