import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt
import pandas as pd
import glob
import configparser
import os

# create empty DataFrame object
df = pd.DataFrame();

# create config
config_ini = configparser.ConfigParser();
config_ini_path = "searchconfig.ini";

if os.path.exists(config_ini_path):
	pass;
else:
	# if .ini file does not exist, ask the user to save the path
	p = input("input the path of the folder where you downloaded .xclx files.");
	f = open(config_ini_path, mode="w");
	f.write("[settings]\nfolder = {}".format(p) + "/*\n");
	f.close();

config_ini = configparser.ConfigParser();
config_ini.read(config_ini_path, encoding='utf-8');
FOLDER = config_ini["settings"]["folder"]

# get a list of files
files = glob.glob(FOLDER)

# extract data
for file in files:
    print("file {} loaded.".format(file));

	# convert into DataFrame object
    records = None;
    if ".xlsx" in file:
    	records = pd.read_excel(file, usecols=[0, 1], index_col=0)
    elif ".csv" in file:
    	records = pd.read_csv(file, usecols=[0, 1], index_col=0);
    else:
    	continue;

    # search
    filtered = records.filter(regex=" 19:15", axis=0);

    try:
    	# discard extra data
    	rec = filtered.iloc[1,:];
    	rec = rec.set_axis([filtered.index.values[1]]);
    	df = pd.concat([df, rec]);

    except:
    	pass;

# organize extracted data
df.index = pd.to_datetime(df.index);
df = df.set_axis(["ViewCount"], axis="columns")
df.sort_index(inplace=True);
print(df);

# display graph
df.to_csv("timetable.csv");
df.plot(title="Total ViewCount in 24 hour");
plt.grid()
plt.show();
