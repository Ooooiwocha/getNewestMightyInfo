from __future__ import print_function
from gspread_formatting import *

import requests
import json
import sys
import os
import time
import quickstart
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import os.path

import gspread
from oauth2client.service_account import ServiceAccountCredentials
hist = 0;
current_rows = 0;
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
def main():
    global hist, current_rows
    if input("PRESS ENTER") != 0:
        URL = 'https://www.googleapis.com/youtube/v3/'

        #ここにプレイリストIDを入力
        PLAYLISTID = "";
        #ここに動画IDを入力 <入力時こちらが優先>
        VIDEOID = "";
        
        # ここにAPI KEYを入力
        API_KEY = ''

        #スプレッドシートを格納するGoogle DriveフォルダのIDを入力
        WORKSPACE = ""
        
        try:
            assert API_KEY != '';
        except:
            print("ERROR: OPEN CODE AND SET YOUTUBE API KEY");
            input("PRESS ENTER TO EXIT");
            sys.exit();
        creds = "None";
        
        try:
            creds = Credentials.from_authorized_user_file('token.json', scope)
        except:
            input("AN ERROR OCCURRED: PUT VALID AUTHENTIFICATED token.json FILE GENERATED AFTER RUNNING quickstart.py BY GOOGLE (MUTATIS MUTANDIS in LIST NAMED \"SCOPE\")");
            sys.exit();

        try:
            assert WORKSPACE != "";
        except:
            input("AN ERROR OCCURRED: OPEN CODE AND INPUT WORKSPACE");
            sys.exit();
        
        gc = gspread.authorize(creds)
        # If there are no (valid) credentials available, let the user log in.

        service = build('drive', 'v3', credentials=creds)

        params = {'key': API_KEY, 'part': 'snippet', 'playlistId': PLAYLISTID,}
        response = requests.get(URL + 'playlistItems', params=params);

        try:
            assert "error" not in response;
        except:
            input("AN ERROR OCCURRED")
            exit();

        video_id = response.json()["items"][0]["snippet"]["resourceId"]["videoId"];
        if VIDEOID != "":
            video_id = VIDEOID;
            
        #編集用シートが作成済みか否かを探索、なければ作成
        results = service.files().list(pageSize=10, fields="nextPageToken, files(id, name)", q="'{}' in parents".format(WORKSPACE) ).execute()

        edit_fileid = "";
        for e in results["files"]:
            if e["name"] == video_id:
                edit_fileid = e["id"];          

        if edit_fileid == "":
            service2 = build('sheets', 'v4', credentials=creds)
            spreadsheet = {
                'properties': {
                    'title': video_id
                }
            }
            spreadsheet = service2.spreadsheets().create(body=spreadsheet, fields='spreadsheetId').execute();
            
            edit_fileid = spreadsheet["spreadsheetId"];
            permission = {
                'type': 'anyone',
                'role': 'commenter',
            }
            service.files().update(fileId=edit_fileid, addParents=WORKSPACE, removeParents=None, fields='id, parents').execute()
            service.permissions().create(fileId=edit_fileid, body=permission).execute()
            print("File Successfully Created.")

        client = gspread.oauth();
        sheet = client.open_by_key(edit_fileid).sheet1
        current_rows = len(sheet.col_values(2));
        hist = sheet.get("B"+str(current_rows))[0][0] if current_rows!=0 else 0;
        
        if sheet.get("A1").first() == None:
            sheet.insert_row(["時刻", "再生数", "高評価数", "コメント数"]);
            current_rows+= 1;
        
        def getResult():
            global hist, current_rows
            
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

                sheet.append_row(arr);
                current_rows+= 1;
                
                for t, e in zip(txt, arr):
                    print(t, e, end=" ");
                    
                hist = int(hist);
                viewCount = int(viewCount);

                if hist > viewCount:
                    print("再生数の減少を検知しました。")
                    fmt = cellFormat(
                        backgroundColor = color(1, 0, 0)
                        )
                    format_cell_range(sheet, "B{}".format(current_rows), fmt);
                                    
                hist = viewCount;                
                
                print();
                print("Data Succcesfully Saved to https://docs.google.com/spreadsheets/d/{}".format(edit_fileid));
                
                f.close();
            except Exception as e:
                print(e)
                print("an error occured");
                f.close();
                sys.exit();
        while True:
            getResult();
            SEC = 60;
            while SEC:
                print("\r{0} ".format(SEC), end="");
                time.sleep(0.98);
                SEC-= 1;
            print("\r", end="");

if __name__ == "__main__":
    main();
