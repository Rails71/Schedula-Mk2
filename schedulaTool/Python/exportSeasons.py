# export seasons
# Wrapper arround the schedula too to make a "clickable" tool to export seasons
import traceback
import sys
import schedulaMain

if __name__ == '__main__':
    try:
        print("Schedula Export tool. Enter a season to export")
        season = input('season: ')
        fileName = season + ".csv"

        print("Exporting season " + season + " to file " + fileName)

        # create the arguments
        argv = []
        argv.append('pullAll')
        argv.append('-s')
        argv.append(season)
        argv.append('-f')
        argv.append(fileName)

        schedulaMain.main(argv)

    except:
        try:
            exc_info = sys.exc_info()
            traceback.print_exception(*exc_info)
            print("Error occured")
            
        finally:
            del exc_info

    finally:
        s = input("Press enter to exit.")
