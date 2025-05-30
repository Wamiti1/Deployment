#API for connecting to MS SQL Server database
from flask import jsonify,Flask,make_response
from flask_restful import Resource, Api, request
from flask_cors import CORS
import pyodbc
from datetime import datetime, time , timedelta
from reportlab.lib.pagesizes import A3,landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
from reportlab.lib import colors
import io
from typing import Optional
import oracledb



#instantiate the flask application
app = Flask(__name__)
CORS(app, origins=["http://localhost:5173"])  # Enable CORS for all routes
#Create an instance of the API class and associate it with the Flask app
api = Api(app)
#Create get_connection() that returns the connection object



def get_connection(
    username: str = "C##Admin",
    password: str = "12345",
    host: str = "localhost",
    port: str = "1521",
    service_name: str = "XEPDB1",
    *,
    thick_mode: bool = False,
    lib_dir: Optional[str] = None
) -> oracledb.Connection:
   
    # Configure thick mode if requested
    if thick_mode:
        try:
            oracledb.init_oracle_client(lib_dir=lib_dir)
        except Exception as e:
            raise RuntimeError(f"Oracle Client initialization failed: {e}") from e
    
    # Build DSN (Data Source Name)
    dsn = f"{host}:{port}/{service_name}"
    
    try:
        # Establish connection in thin mode (default)
        connection = oracledb.connect(
            user=username,
            password=password,
            dsn=dsn,
            # Enable statement caching
            stmtcachesize=20
        )
                
        return connection
        
    except oracledb.DatabaseError as e:
        error, = e.args
        if error.code == 1017:  # Invalid credentials
            raise oracledb.DatabaseError("Authentication failed - check username/password") from e
        elif error.code == 12541:  # Listener error
            raise oracledb.DatabaseError(f"Cannot connect to {dsn} - check host/port/service_name") from e
        else:
            raise oracledb.DatabaseError(f"Database connection failed: {error.message}") from e
        
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
                "ALUMNI_INFO" : "SELECT * FROM ALUMNI_INFO",
                "AWARDS" : "SELECT * FROM AWARDS",
                "CHAPTERS"  : "SELECT * FROM CHAPTERS" ,
                "EVENTREGISTRATION" : "SELECT * FROM EVENTREGISTRATION",
                "ALUMNIOFFICE"  : "SELECT * FROM ALUMNIOFFICE",
                "EVENTS"  : "SELECT * FROM ALLEVENTS",
                "OTHERINSTITUTIONS"  : "SELECT * FROM OTHERINSTITUTIONS",
     
            }       
            
                   
            if tableName not in sql : 
                return jsonify({"message" : "The table does not exist"}) 
            cursor.execute(sql[tableName]) 
            
            allFetched = cursor.fetchall()
            # Check if any records were fetched
            if not allFetched:
                return jsonify({"message": "No records found"})
            columns = [c[0] for c in cursor.description]
            data = [list(row) for row in allFetched]
                
            
            for i in range(len(data)) :
                data[i] = dict(zip(columns, data[i]))
            
            response = make_response(jsonify(data))
            response.headers['Content-Type'] = 'application/json'
            response.headers['Cache-Control'] = 'public,max-age=60'
            cursor.close()
            connection.close()
            return response
        
            
        except Exception as e :
            return jsonify({"message" : str(e)})          

    def post(self, tableName):
        try:
            # A JSON Object that shows the tables in the dB and their columns
            # This is used to validate the data being inserted into the database
            tableNames = {
                "ALUMNI_INFO" : ['Alumni_ID', 'Alumni_Name', 'Chapter_ID', 'Phone_Number', 'Graduation_Year', 'Degree', 'Email', 'Industry'],
                "AWARDS" : ['Awards_ID', 'AwardName', 'Recipient_ID', 'Date_Of_Issue', 'AwardsDescription'],
                "CHAPTERS"  : ['Chapter_ID', 'Chapter_Location', 'Chapter_Name', 'Contact_Person_Name', 'Email', 'Established_Year'] ,
                "EVENTREGISTRATION" : ['AlumniID', 'EventID', 'RegistrationDate', 'EmailAddress','PhoneNumber', 'PaymentStatus'],
                "ALUMNIOFFICE"  : ['Office_Name', 'Office_ID', 'Office_Location', 'Contact_Info'],
                "EVENTS"  : [ 'ProgramName','Venue', 'ProgramDate','Start_Time','End_Time', 'Chapter_ID','Contact_Info', 'ProgramDescription'],
                "OTHERINSTITUTIONS"  : ['Institution_Name', 'Institution_ID', 'Location', 'Partnership_Type', 'Contact_Info']
            }
            #Checks if one has provided a table that does not exist in the database 
            if tableName not in tableNames:
                return jsonify({"message": "The table does not exist"})

            data = request.json
            # Check if data is empty
            if not data:
                return jsonify({"message": "No data provided"})

            # Validate all required fields are present
            required_fields = tableNames[tableName]
            missing_fields = [field for field in required_fields if field not in data]
            if missing_fields:
                return jsonify({
                    "message": "Missing required fields",
                    "missing_fields": missing_fields
                })

            # Generate placeholders based on number of fields
            placeholders =  ', '.join([':%s' % (i+1) for i in range(len(required_fields))])  # Oracle style
            columns = ', '.join(required_fields)
            
            # Use parameterized query for safety
            sql = f"INSERT INTO {tableName} ({columns}) VALUES ({placeholders})"
            values = tuple(data[field] for field in required_fields)

            with get_connection() as connection:
                with connection.cursor() as cursor:
                    cursor.execute(sql, values)
                    connection.commit()
                   
            return jsonify({"message": "Data inserted successfully"})

        except Exception as e:
            return jsonify({"message": str(e)})
        
api.add_resource(Tables, '/table/<string:tableName>')
  
class Views(Resource) :
    def get(self, viewsName) :
        try : 
            connection = get_connection()
            cursor = connection.cursor()
            sql = {
                'ALUMNISBETWEEN2005AND2015' : "SELECT * FROM ALUMNISBETWEEN2005AND2015",
                
                'InstitutionContactSummary'  : "SELECT * FROM InstitutionContactSummary",
                
                'AWARDSBETWEEN2020AND2022' : "SELECT * FROM AWARDSBETWEEN2020AND2022",
                
                'MERUANDNAIROBICHAPTERSALUMNIS' : "SELECT * FROM MERUANDNAIROBICHAPTERSALUMNIS",
                
                'OTHERINSTITUTIONSVIEW' : "SELECT * FROM OTHERINSTITUTIONSVIEW",
                
                'TECHNOLOGYALUMNIS' : "SELECT * FROM TECHNOLOGYALUMNIS",
                
                'UPCOMINGEVENTS'  : "SELECT  * FROM UPCOMINGEVENTS"
            }
            
            if viewsName not in sql : 
                return jsonify({"message" : "The view does not exist"})
            cursor.execute(sql[viewsName])
            
            
            allFetched = cursor.fetchall()
            # Check if any records were fetched
            if not allFetched:
                return jsonify({"message": "No records found"})
            columns = [c[0] for c in cursor.description]
            data = [list(row) for row in allFetched]
                    
                
            for i in range(len(data)) :
                    data[i] = dict(zip(columns, data[i]))
                
            
                
            response = make_response(jsonify(data))
            response.headers['Content-Type'] = 'application/json'
            response.headers['Cache-Control'] = 'public,max-age=60'
            cursor.close()
            connection.close()
            return response
        except Exception as e :
            return jsonify({"message" : str(e)})       
#Create an endpoint for views
api.add_resource(Views,'/view/<string:viewsName>')

#TODO : Modify the Reports Class to dynamically generate reports
class Reports(Resource):
    def get(self, reportName):
        try:
            connection = get_connection()
            cursor = connection.cursor()
            sql = {
                'ALUMNISBETWEEN2005AND2015' : "SELECT * FROM ALUMNISBETWEEN2005AND2015",
                
                'ALUMNIDIRECTORY'  : "SELECT * FROM ALUMNIDIRECTORY",
                
                'AWARDSBETWEEN2020AND2022' : "SELECT * FROM AWARDSBETWEEN2020AND2022",
                
                'MERUANDNAIROBICHAPTERSALUMNIS' : "SELECT * FROM MERUANDNAIROBICHAPTERSALUMNIS",
                
                'OTHERINSTITUTIONSVIEW' : "SELECT * FROM OTHERINSTITUTIONSVIEW",
                
                'TECHNOLOGYALUMNIS' : "SELECT * FROM TECHNOLOGYALUMNIS",
                
                'UPCOMINGEVENTS'  : "SELECT  * FROM UPCOMINGEVENTS"
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
            elements.append(Paragraph(f"{reportName} {datetime.today().strftime('%B %d, %Y')} {(datetime.now()).strftime('%I:%M %p')}", title_style))

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
api.add_resource(Reports, '/report/<string:reportName>')
                                


if __name__ == '__main__':
    app.run(debug=True, host="0.0.0.0")

# if __name__ == '__main__':
#     app.run()
       