import requests
import json
import pandas as pd

# All category list
categories = ['ai-machine-learning',
              'analytics',
              'collaboration',
              'commerce',
              'compliance-legals',
              'customer-service',
              'finance',
              'geolocation',
              'human-resources',
              'internet-of-things',
              'it-management-tools',
              'marketing',
              'operations',
              'productivity',
              'project-management',
              'sales']


def getAppDetails(page_number, category):
   """
   Get the app data for a single page.
   
   Args:
         page_number (int): The page number that should be crawled.
         category (str): The targeted category.

   Returns:
         A list having a list of app_details data
         and a integer value of app_data_count.
   """

   headers = {"Content-Type":"application/json",
              "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"}
   paramsGet = {"country":"US",
                "page":f"{page_number}",
                "ReviewsMyCommentsFilter":"true",
                "category":f"{category}",
                "region":"ALL",
                "entityType":"App"}
   base_url = 'https://appsource.microsoft.com/view/tiledata'
   response = requests.get(base_url, params=paramsGet, headers=headers)

   jsonData = json.loads(response.content)
   appsData = jsonData['apps']['dataList']

   app_data_count = jsonData['apps']['count']

   app_details = []
   if len(appsData)>4:
      for data in appsData[4:]:
         TotalRatings = None
         AverageRating = None
         for r in data['ratingSummaries']:
            if r['Source']=='All':
               TotalRatings = r['TotalRatings']
               AverageRating = r['AverageRating']
               break

         categories_list = []
         for c in data['categoriesDetails']:
            categories_list.append(c['longTitle'])
               
         app_details.append({
            'title': data['title'],
            'publisher': data['publisher'],
            'builtFor': data['builtFor'],
            'actionString': data['actionString'],
            'total_ratings': TotalRatings,
            'average_rating': AverageRating,
            'belong_categories': categories_list,
            })
   return [app_details, app_data_count]
            

if __name__=='__main__':

   # Iterating through all the categories one by one
   for category in categories:
      print(f'Scraping for {category} category.')
      
      all_final_data = []
      page_number = 1
      while True:
         print(f'Page number = {page_number}')
         
         app_data_list = getAppDetails(page_number, category)
         
         for app_data in app_data_list[0]:
            all_final_data.append(app_data)
         if page_number*60 >= app_data_list[1]:
            break
         page_number+=1

      # Remove duplicates and build new data list
      new_final_data = []
      for final_data in all_final_data:
         if final_data not in new_final_data:
            new_final_data.append(final_data)

      # Saving file in Excel format with using pandas dataframe
      df = pd.DataFrame(new_final_data)
      df.to_excel(f'{category}.xlsx', index=False)
