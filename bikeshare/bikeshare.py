import time

import pandas as pd

import numpy as np



CITY_DATA = { 'chicago': 'chicago.csv',

              'new york city': 'new_york_city.csv',

              'washington': 'washington.csv' }



def get_filters():

    """

    Asks user to specify a city, month, and day to analyze.



    Returns:

        (str) city - name of the city to analyze

        (str) month - name of the month to filter by, or "all" to apply no month filter

        (str) day - name of the day of week to filter by, or "all" to apply no day filter

    """

    print('Hello! Let\'s explore some US bikeshare data!')

    # TO DO: get user input for city (chicago, new york city, washington). HINT: Use a while loop to handle invalid inputs

    city = ''

    while city.lower() not in ['chicago', 'new york city', 'washington']:

       city = input('What city do you want? The choices are Chicago, New York City, or Washington.  ')

       city = city.lower()

        

# TO DO: get user input for month (all, january, february, ... , june)

    month = ''

    while month.lower() not in ['all', 'january', 'february', 'march', 'april', 'may','june']: 

       month = input('What month do you want? The choices are: all, january, february, march, april, may, or june.  ')

       month = month.lower()



    # TO DO: get user input for day of week (all, monday, tuesday, ... sunday)

    day = ''

    while day.lower() not in ['all', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday','saturday', 'sunday']: 

       day = input('What day do you want? The choices are: all, monday, tuesday, ... , sunday.  ')

       day = day.lower()





    print('-'*40)

    return city, month, day





def load_data(city, month, day):

    """

    Loads data for the specified city and filters by month and day if applicable.



    Args:

        (str) city - name of the city to analyze

        (str) month - name of the month to filter by, or "all" to apply no month filter

        (str) day - name of the day of week to filter by, or "all" to apply no day filter

    Returns:

        df - Pandas DataFrame containing city data filtered by month and day

    """

    # load data file into dataframe

    df = pd.read_csv(CITY_DATA[city])



    # convert the Start Time column to datetime

    df['Start Time'] = pd.to_datetime(df['Start Time'])



    # extract month and day of week from Start Time to create new columns

    df['month'] = df['Start Time'].dt.month

    df['day_of_week'] = df['Start Time'].dt.weekday_name



    # filter by month if applicable

    if month != 'all':

        # use the index of the months list to get the corresponding int

        months = ['january', 'february', 'march', 'april', 'may', 'june']

        month = months.index(month) + 1



        # filter by month to create the new dataframe

        df = df[df['month'] == month]



    # filter by day of week if applicable

    if day != 'all':

        # filter by day of week to create the new dataframe

        df = df[df['day_of_week'] == day.title()]



    return df





def time_stats(df):

    """Displays statistics on the most frequent times of travel."""



    print('\nCalculating The Most Frequent Times of Travel...\n')

    start_time = time.time()

    

    df['Start Time'] = pd.to_datetime(df['Start Time'])





    # TO DO: display the most common month

    df['month'] = df['Start Time'].dt.month



    popular_month = df['month'].value_counts().idxmax()

    





    # TO DO: display the most common day of week

    df['week'] = df['Start Time'].dt.week



    popular_week = df['week'].value_counts().idxmax()





    # TO DO: display the most common start hour



# extract hour from the Start Time column to create an hour column

    df['hour'] = df['Start Time'].dt.hour



# find the most common hour (from 0 to 23)

    popular_hour = df['hour'].value_counts().idxmax()

    

    print('Most frequent month:', popular_month)

    print('Most common Week:', popular_week)

    print('Most Frequent Start Hour:', popular_hour)



    print("\nThis took %s seconds." % (time.time() - start_time))

    print('-'*40)





def station_stats(df):

    """Displays statistics on the most popular stations and trip."""



    print('\nCalculating The Most Popular Stations and Trip...\n')

    start_time = time.time()



    # TO DO: display most commonly used start station

    

    common_start = df['Start Station'].value_counts().idxmax()





    # TO DO: display most commonly used end station

    common_end = df['End Station'].value_counts().idxmax()



    # TO DO: display most frequent combination of start station and end station trip

    

    common_start_end = (df['Start Station'] + " >> " + df['End Station']).value_counts().idxmax()

    

    print('Most common starting station:', common_start)

    print('Most common ending station:', common_end)

    print('Most common combination 0f start and ending stations:', common_start_end)



    

    



    print("\nThis took %s seconds." % (time.time() - start_time))

    print('-'*40)





def trip_duration_stats(df):

    """Displays statistics on the total and average trip duration."""



    print('\nCalculating Trip Duration...\n')

    start_time = time.time()



    # TO DO: display total travel time

    total_time = df['Trip Duration'].sum()

    

    print('total trip duration:', total_time)

    



    # TO DO: display mean travel time

    mean = df['Trip Duration'].mean()

    print('average trip time:', mean)



    print("\nThis took %s seconds." % (time.time() - start_time))

    print('-'*40)





def user_stats(df):

    """Displays statistics on bikeshare users."""



    print('\nCalculating User Stats...\n')

    start_time = time.time()



    # TO DO: Display counts of user types

    

    user_types = df['User Type'].value_counts()



    print(user_types)

    # TO DO: Display counts of gender

    if  'Gender' in df.columns:

       gender = df['Gender'].value_counts()

       print(gender)

    else:

       print('no gender data for washington')





    # TO DO: Display earliest, most recent, and most common year of birth

    if  'Birth Year' in df.columns:

       earliest_birth = df['Birth Year'].min()

       recent_birth = df['Birth Year'].max()

       common_birth = df['Birth Year'].value_counts().idxmax()

    

       print('earliest birth year:', earliest_birth)

       print('most recent birth year:', recent_birth)

       print('most common birth year:', common_birth)

    else:

       print('no birth year data for washington')

    





    print("\nThis took %s seconds." % (time.time() - start_time))

    print('-'*40)



def raw_data(df):

    x = 0

    y = 5

    response = input('display 5 rows of raw data? yes or no?  ')

    while response.lower() == 'yes':

       print(df.iloc[x:y])

       x += 5

       y += 5

       response = input('display 5 rows of raw data? yes or no?  ')

       



    

def main():

    while True:

        city, month, day = get_filters()

        df = load_data(city, month, day)



        time_stats(df)

        station_stats(df)

        trip_duration_stats(df)

        user_stats(df)

        raw_data(df)



        restart = input('\nWould you like to restart? Enter yes or no.\n')

        if restart.lower() != 'yes':

            break





if __name__ == "__main__":

	main()
