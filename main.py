#Briana DeAndrade
#Nothing in the required directions (not including Extra Credit) was left undone.
#In order to run the program correctly, please see the corresponding numbers next to each command

#I imported openpyxl and numbers since we are working with Excel and doing math later in the code
import openpyxl
import numbers

#Here in my main function, I opened my worksheet, created a string of acceptable answers and prompt the user the question as to what they would like to see
def main():
    game_worksheet = open_worksheet("games-features.xlsx")
    run_fxn_answers = ["1", "2", "3", "4", "5", "6"]
    while True:
        response = input("What command would you like to see? Type the corresponding number next to the command you'd like.\n"
          "0. Program quit? 1. Find the most players? 2.Find age restricted games? \n"
          "3. Find often recommended games? 4.Find well played games? \n"
          "5. See a count for the number of games? Or, 6. Look up more data?\n")
        response = response.lower()
#Depending on what the user wants, the following if statements guide the code to the appropriate function(s) and pass the worksheet as a parameter for it to run correctly
        if response not in run_fxn_answers: break

        if response == "1":
            find_Most_Players(game_worksheet)

        if response == "2":
            find_Age_Restricted(game_worksheet)

        if response == "3":
            find_Often_Recommended(game_worksheet)

        if response == "4":
            find_well_played(game_worksheet)

        if response == "5":
            find_Count_For_System(game_worksheet)

        if response == "6":
            Look_Up_Data(game_worksheet)

#In my open_worksheet function, I pass the file, activate the sheet and return it so we can use it throughout all of the following functions
def open_worksheet(file):
    game_excel = openpyxl.load_workbook(file)
    game_sheet = game_excel.active
    return game_sheet

#In the find_Most_Players function I first let the user know that they chose this.
#Then I ask the user what platform they'd like to see the most players from.
def find_Most_Players(game_worksheet):
    largest_row = None
    print("You chose to Find Most Players")
    fmp_inp = input("Which system, Windows, Mac, or Linux, do you want to consider?")
    fmp_inp = fmp_inp.lower()
#There are three if statements depending on what platform the user wants to see.
    #In each, there is a for loop that checks to see where the largest row is within the columns that have "True" as being a platform for that specific game
    #From there, each for loop prints the Game Title, Release Date and then returns the command.
    if fmp_inp == "windows":
        for row in game_worksheet.rows:
            windows_platform = row[26]
            windows_platform_value = windows_platform.value
            if windows_platform_value == "True":
                if largest_row is None:
                    largest_row = row
                if largest_row[17].value < row[17].value:
                    largest_row = row
        print(f"Windows: Game title: {largest_row[2].value}, Release date: {largest_row[4].value}, Player Count: {largest_row[17].value}")
        return
    if fmp_inp == "mac":
        for row in game_worksheet.rows:
            mac_platform = row[28]
            mac_platform_value = mac_platform.value
            if mac_platform_value == "True":
                if largest_row is None:
                    largest_row = row
                if largest_row[17].value < row[17].value:
                    largest_row = row
        print(f"Mac: Game title: {largest_row[2].value}, Release date: {largest_row[4].value}, Player Count: {largest_row[17].value}")
        return
    if fmp_inp == "linux":
        for row in game_worksheet.rows:
                linux_platform = row[27]
                linux_platform_value = linux_platform.value
                if linux_platform_value == "True":
                    if largest_row is None:
                        largest_row = row
                    if largest_row[17].value < row[17].value:
                        largest_row = row
        print(f"Windows: Game title: {largest_row[2].value}, Release date: {largest_row[4].value}, Player Count: {largest_row[17].value}")
        return

#In the find_Age_Restricted fxn, I first let the user know they've selected this function.
# I established a for loop, created variables, and then printed the age restricted rows and date of release as asked.
def find_Age_Restricted(game_worksheet):
    print("You chose to Find Age Restricted Games")
    for row in game_worksheet.rows:
        age_restrictions = row[5].value
        response_name = row[3].value
        release_date = row[4].value

        if not isinstance(age_restrictions, numbers.Number):
            continue
        if age_restrictions >= 17:
            print(f"Name:{response_name}, Date of release: {release_date}")

#In the find_Often_Recommended fxn, I first let the user know they've selected this function.
#I then ask the user what is the minimum number of recommendations they are looking for. From there I established a for loop that did some math to also incorporate the percent of recommendations/owner
#The recommendation count needs to be a number and the number of owners cannot be 0 seeing as we will be dividing by that. After, the statement containing all info will be presented to user.
def find_Often_Recommended(game_worksheet):
    print("You chose to Find Recommended Games")
    for_inp = int(input("What is the minimum number of recommendations you're looking for?"))
    for row in game_worksheet.rows:
        recommendation_count = row[12].value
        number_of_owners = row[15].value
        if not isinstance(recommendation_count, numbers.Number):
            continue
        if number_of_owners == 0:
            continue
        recommendation_percent = recommendation_count / number_of_owners
        game_name = row[2].value

        if recommendation_count > for_inp:
            print(f"Name:{game_name}, Recommendation Count:{recommendation_count}, Number of owners:{number_of_owners}, Percent of Recommendations:{recommendation_percent}")

#In the find_well_played fxn, I first let the user know they've selected this function. I then ask them for a cutoff percentage.
#A for loop is establised and a similar kind of work is done as the previous function. The percentage and title are then presented in a statement to the User.
def find_well_played(game_worksheet):
    print("You chose to Find Well Played Games")
    fwp_answer = int(input("What should the cutoff percentage be?"))

    for row in game_worksheet.rows:
        num_of_players = row[17].value
        num_of_owners = row[15].value
        if not isinstance(num_of_players, numbers.Number):
            continue
        if num_of_owners == 0:
            continue
        fwp_percent = num_of_players / num_of_owners
        game_title = row[2].value
        if fwp_percent > fwp_answer:
            print(f"Percentage: {fwp_percent}, Game Title: {game_title}")

#In the find_Count_For_System fxn, I first let the user know they've selected this function.
#We establish that the count should start from 0 and then a for loop
#Within the for loop, if the platform presents itself to be true, we count or "+1" to the count. From there a count of games on all platforms is counted.
def find_Count_For_System(game_worksheet):
    print("You chose to see a count for the number of games")
    num_platforms = {"windows": 0, "mac": 0, "linux": 0}
    for row in game_worksheet.rows:
        mac_platform = row[28].value
        windows_platform = row[26].value
        linux_platform = row[27].value
        if windows_platform == "True":
            num_platforms["windows"] +=1

        if mac_platform == "True":
            num_platforms["mac"] +=1

        if linux_platform == "True":
            num_platforms["linux"] +=1
    print(f"Windows Platform: {num_platforms['windows']}")
    print(f"Mac Platform: {num_platforms['mac']}")
    print(f"Linux Platform: {num_platforms['linux']}")
    return num_platforms

#In the Look_Up_Data fxn, I first let the user know they've selected this function. I then ask the User which specific game they'd like more data on
#We establish a for loop, and then a few if statements where if the platform reads "True" for a certain game, it will then print that out in a final statement
#The final statement includes the game name, platform it runs on, metacritic score, DLC count and owner count. Then the function is returned.
def Look_Up_Data(game_worksheet):
    print("You chose to look up more data")
    data_question = input("Which game would you like more data on?")
    data_question = data_question.lower()
    for row in game_worksheet.rows:
        runs_on = "Runs on: "
        mac_platform = row[28].value
        windows_platform = row[26].value
        linux_platform = row[27].value
        response_name = row[3].value
        metacritic_score = row[9].value
        DLC_Count = row[8].value
        owner_count = row[15].value
        if isinstance(response_name, numbers.Number):
            continue
        if mac_platform == "True":
            runs_on = runs_on + "mac"
        if windows_platform == "True":
            runs_on = runs_on + "windows"
        if linux_platform == "True":
            runs_on = runs_on + "linux"
        if response_name.lower() == data_question.lower():
            print(f"Name: {response_name}, {runs_on} Metacritic Score: {metacritic_score}, DLC Count: {DLC_Count}, Owner Count: {owner_count}" )
            return


#Main is called down here at the end and the project meets all direction criteria. :)
main()

