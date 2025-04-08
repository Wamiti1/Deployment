#API for connecting to MS SQL Server database
from flask import jsonify,Flask,make_response
from flask_restful import Resource, Api
from flask_cors import CORS
import pyodbc
from datetime import datetime, time
from reportlab.lib.pagesizes import A3,landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
from reportlab.lib import colors
import io
import pymssql


#instantiate the flask application
app = Flask(__name__)
CORS(app, origins=["http://localhost:5173"])  # Enable CORS for all routes
#Create an instance of the API class and associate it with the Flask app
api = Api(app)
#Create get_connection() that returns the connection object
def get_connection():
    connection = pymssql.connect(
        server='alumnioffice.database.windows.net',
        user='group31',
        password='train12$',
        database='UniversityAlumniOffice',
        port=1433,  # Ensure you're using the default SQL Server port
        timeout=30  # Set a timeout for the connection
    )
    return connection
# Helper function to serialize time objects
def serialize_time(obj):
    if isinstance(obj,time):
        return obj.isoformat()  # Convert time object to ISO format string
    raise TypeError("Type not serializable")
            
class Tables(Resource) : 
    def get(self, tableName)  :
        try : 
            connection = get_connection()
            cursor = connection.cursor()
            
            sql = {
                'Alumni_Info' : "SELECT * FROM Alumni_Info",
                'Awards' : "SELECT * FROM Awards",
                'Chapters'  : "SELECT * FROM Chapters" ,
                'EventRegistration' : """
                SELECT 
                AttendanceID,
                AlumniID,
                AlumniName,
                EventID,
                CAST(RegistrationDate AS NVARCHAR) AS RegistrationDate,
                EmailAddress,
                PhoneNumber,
                PaymentStatus,
                Attended,
                CheckInTime,
                CheckOutTime,
                Comments
                FROM EventRegistration
                """,
                'AlumniOffice'  : "SELECT * FROM AlumniOffice",
                'Events'  : "SELECT * FROM Events",
                'OtherInstitutions'  : "SELECT * FROM OtherInstitutions",
            }       
            
            
            
            if tableName not in sql : 
                return jsonify({"message" : "The table does not exist"}) 
            cursor.execute(sql[tableName]) 
            
            if cursor.rowcount == 0 :
                return jsonify({'Message' :  'No records found'})  
                        
            
            allFetched = cursor.fetchall()
            columns = [c[0] for c in cursor.description]
            data = [list(row) for row in allFetched]
                
            
            for i in range(len(data)) :
                data[i] = dict(zip(columns, data[i]))
            
            if tableName == 'EventRegistration' :
                for i in range(len(data)) :
                    data[i]['CheckInTime'] = serialize_time(data[i]['CheckInTime']) if data[i]['CheckInTime'] else None
                    data[i]['CheckOutTime'] = serialize_time(data[i]['CheckOutTime']) if data[i]['CheckOutTime'] else None
                
            
            response = make_response(jsonify(data))
            response.headers['Content-Type'] = 'application/json'
            response.headers['Cache-Control'] = 'public,max-age=60'
            cursor.close()
            connection.close()
            return response
        
            
        except Exception as e :
            return jsonify({"message" : str(e)})
       
           
api.add_resource(Tables, '/table/<string:tableName>')
  
class Views(Resource) :
    def get(self, viewsName) :
        try : 
            connection = get_connection()
            cursor = connection.cursor()
            sql = {
                'AlumnisBetween2005and2015' : """SELECT * FROM AlumnisBetween2005and2015""",
                
                'AlumniDirectory'  : """SELECT * FROM AlumniDirectory""",
                
                'awardsBetween2020and2022' : """SELECT * FROM AwardsBetween2020and2022""",
                
                'meruAndNairobiChapters' : """SELECT * FROM [MeruandNairobiChapters'Alumnis]""",
                
                'otherInstitutions' : """SELECT * FROM [OtherInstitutionsView]""",
                
                'TechnologyAlumnis' : """SELECT * FROM [TechnologyAlumnis]""",
                
                'upcomingEvents'  : """SELECT  * FROM [UpcomingEvents]""",
            }
            
            if viewsName not in sql : 
                return jsonify({"message" : "The view does not exist"})
            cursor.execute(sql[viewsName])
            
            if cursor.rowcount == 0 : 
                return jsonify({"message" : "No records found"})
            else :
                allFetched = cursor.fetchall()
                columns = [c[0] for c in cursor.description]
                data = [list(row) for row in allFetched]
                    
                
                for i in range(len(data)) :
                    data[i] = dict(zip(columns, data[i]))
                
                if viewsName == 'upcomingEvents' :
                    for i in range(len(data)) :
                        data[i]['CheckInTime'] = serialize_time(data[i]['CheckInTime']) if data[i]['CheckInTime'] else None
                        data[i]['CheckOutTime'] = serialize_time(data[i]['CheckOutTime']) if data[i]['CheckOutTime'] else None
                    
                
                response = make_response(jsonify(data))
                response.headers['Content-Type'] = 'application/json'
                response.headers['Cache-Control'] = 'public,max-age=60'
                cursor.close()
                connection.close()
                return response
        except Exception as e :
            return jsonify({"message" : str(e)})
        finally :
            cursor.close()
            connection.close()
#Create an endpoint for views
api.add_resource(Views,'/views/<string:viewsName>')

#TODO : Modify the Reports Class to dynamically generate reports
class Reports(Resource):
    def get(self, reportName):
        try:
            connection = get_connection()
            cursor = connection.cursor()
            sql = {
                'AlumnisBetween2005and2015': """SELECT * FROM AlumnisBetween2005and2015""",
                'AlumniDirectory': """SELECT * FROM AlumniDirectory""",
                'awardsBetween2020and2022': """SELECT * FROM AwardsBetween2020and2022""",
                'meruAndNairobiChapters': """SELECT * FROM [MeruandNairobiChapters'Alumnis]""",
                'otherInstitutions': """SELECT * FROM [OtherInstitutionsView]""",
                'TechnologyAlumnis': """SELECT * FROM [TechnologyAlumnis]""",
                'upcomingEvents': """SELECT * FROM [UpcomingEvents]""",
            }

            if reportName not in sql:
                return jsonify({'message': 'The report does not exist'})

            cursor.execute(sql[reportName])
            columns = [column[0] for column in cursor.description]
            data = [columns] + [list(row) for row in cursor.fetchall()]

            if len(data) <= 1:
                return jsonify({"message": "No records found; thus PDF cannot be generated"})

            buffer = io.BytesIO()
            doc = SimpleDocTemplate(
                buffer,
                pagesize=landscape(A3),
                title=f"{reportName} (Digital Use)"
            )

            elements = []
            styles = getSampleStyleSheet()

            # Title Style
            title_style = ParagraphStyle(
                'TitleStyle',
                parent=styles['Title'],
                fontSize=24,
                spaceAfter=20,
                textColor=colors.darkcyan
            )
            elements.append(Paragraph(f"{reportName} {datetime.now()}", title_style))

            # Table Style
            table_data = []
            for i, row in enumerate(data):
                if i == 0:
                    # Header row
                    table_data.append([Paragraph(col, styles['Heading2']) for col in row])
                else:
                    # Alternating row colors
                    row_color = colors.beige if i % 2 == 0 else colors.white
                    table_data.append([Paragraph(str(cell), styles['BodyText']) for cell in row])

            table = Table(table_data)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 14),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), row_color),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('LEFTPADDING', (0, 0), (-1, -1), 10),
                ('RIGHTPADDING', (0, 0), (-1, -1), 10),
                ('TOPPADDING', (0, 0), (-1, -1), 5),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ]))

            elements.append(table)
            doc.build(elements)

            buffer.seek(0)
            response = make_response(buffer.getvalue())
            response.headers['Content-Type'] = 'application/pdf'
            response.headers['Content-Disposition'] = f'attachment; filename={reportName}.pdf'
            cursor.close()
            connection.close()
            return response

        except Exception as e:
            return jsonify({"message": str(e)})

        
# Add the resource to the API
api.add_resource(Reports, '/reports/<string:reportName>')
                                


# if __name__ == '__main__':
#     app.run(debug=True, host="0.0.0.0")

if __name__ == '__main__':
    app.run()
       