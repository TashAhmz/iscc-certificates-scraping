<<<<<<< HEAD
t = [1, 2, 3, 4, 5, 6, 7]
print(t[1:-1])
print("hello world".capitalize())
=======
import pycities as c

db = c.CityDatabase(fetch_fields=("id", "name", "administrative_name", "country_name", "longitude", "latitude"))
db.connect()
print(db.search(query="manlius", lang="uk", limit=1))
>>>>>>> 3edeeec163bac95bc8636d5d29e4b8fc89a53075
