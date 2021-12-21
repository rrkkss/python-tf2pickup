map_list_single = []
map_list_all = []

def MapsCount(js):
    for n in js:
        if n == 'info':
            map_info = list(js[n].values())
            print(map_info[0])
            if map_info[0] not in map_list_single:
                map_list_single.append(map_info[0]) #create list of uniques
            map_list_all.append(map_info[0])        #throw them all in for comparison later

def MapsPrint():
    n = 0
    for i in map_list_single:
        print(map_list_single[n],': #',map_list_all.count(i))
        n += 1