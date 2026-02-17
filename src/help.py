import pycities as c

db = c.CityDatabase(fetch_fields=("id", "name", "administrative_name", "country_name", "longitude", "latitude"))
db.connect()
print(db.search(query="manlius", lang="uk", limit=1))