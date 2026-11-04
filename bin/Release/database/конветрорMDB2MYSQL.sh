for t in $(mdb-tables -1 basekomp.mdb); do
  echo "-- Table: $t" >> data.sql
  mdb-export -I mysql basekomp.mdb "$t" >> data.sql
done

cat schema.sql data.sql > full_database.sql