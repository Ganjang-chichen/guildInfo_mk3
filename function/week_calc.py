import datetime

def get_weekinfo(week_count) :
    now = datetime.datetime.now()
    weeks = now.isocalendar()

    first = now + datetime.timedelta(-int(weeks[2]) + 1 - (7 * week_count))
    last = now + datetime.timedelta(7 - int(weeks[2]) - (7 * week_count))

    week_info = {
        "first" : {
            "year" : (first.strftime("%Y")),
            "month" : (first.strftime("%m")),
            "day" : (first.strftime("%d"))
        },
        "last" : {
            "year" : (last.strftime("%Y")),
            "month" : (last.strftime("%m")),
            "day" : (last.strftime("%d"))
        }
    }

    return week_info
    
def get_week_range(weekinfo) :
    sheet_name = (  weekinfo["first"]["year"] + weekinfo["first"]["month"] + weekinfo["first"]["day"] + "-" + 
                    weekinfo["last"]["year"] + weekinfo["last"]["month"] + weekinfo["last"]["day"])
    return sheet_name


if __name__ == '__main__' :
    info = get_weekinfo(1)
    print(info)
    print(info.get("fff"))
    d = {"hello" : 1}
   
    if d.get("e") == None:
        print("n")
