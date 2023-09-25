# DominoChangeDatabaseReplicaID
Change Domino database ReplicaId (C Notes API)

See demo.lss (you only need to adjust function: GetDatabase that returns server andf filepath of database).

How to use the library DbUtils

Dim session As NotesSession
Dim database As NotesDatabase
Dim dbUtils As DbUtils
Dim timedate As TIMEDATE
	
Set session = New NotesSession
Set database = session.Getdatabase("hexagon/explicants", "test2.nsf", false)

Set dbUtils = New DbUtils(session)
	
Print "1) setDbReplicaIdRandom"
Print "before: " & database.replicaID
Call dbUtils.setDbReplicaIdRandom(database.server, database.filepath)
Print "after " & database.replicaID
	
Print "2) setDbReplicaIdByTimeDate"
Print "before: " & database.replicaID
timedate.Innards(0) = 1000
timedate.Innards(1) = 64000
Call dbUtils.setDbReplicaIdByTimeDate(database.server, database.filepath, timedate)
Print "after " & database.replicaID
	
Print "3) setDbReplicaIdByTimeDate"
Print "before: " & database.replicaID
Call dbUtils.setDbReplicaIdByString(database.server, database.filepath, "1234567887654321")
Print "after " & database.replicaID
