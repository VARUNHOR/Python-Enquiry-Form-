import pymysql

def create_table():
    connection = pymysql.connect(host="localhost", user="root", passwd="", database="employee")
    cursor = connection.cursor()
    query = """
    CREATE TABLE IF NOT EXISTS enquiry_data (
        id INT AUTO_INCREMENT PRIMARY KEY,
        date VARCHAR(50),
        name VARCHAR(100),
        mobile_no VARCHAR(15),
        alternate_no VARCHAR(15),
        email_id VARCHAR(100),
        address TEXT,
        course_interested VARCHAR(100),
        batch_preferred VARCHAR(100),
        how_you_came_to_know_us TEXT,
        experience_status VARCHAR(50),
        contact_person VARCHAR(100),
        counselor VARCHAR(100),
        fees VARCHAR(20),
        comment TEXT,
        enquiry VARCHAR(10),
        registration VARCHAR(10)
    );
    """
    cursor.execute(query)
    connection.commit()
    print("Table created successfully!")
    connection.close()

create_table()
