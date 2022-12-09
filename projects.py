"""
Class represenation of the projects table
"""
import sqlite3

class Project:
    def __init__(self, project_id, con):
        self.con: sqlite3.Connection
        self.con = con
        self.cur = self.con.cursor()
        self._project_id = project_id

        self.cur.execute(
            f"select project_type, project_title, user_id \
            from projects where project_id={self._project_id}"
            )
        
        results = self.cur.fetchall()
        assert(len(results) == 1)
        (
            self._project_type, self._project_title, self._user_id,
            ) = results[0]

    @property
    def project_type(self):
        return self._project_type
    
    @property
    def project_title(self):
        return self._project_title
    
    @property
    def user_id(self):
        return self._user_id
    
    @property
    def project_id(self):
        return self._project_id
