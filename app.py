import uuid # random id generator
from streamlit_option_menu import option_menu
from settings import *
import os
import time
import win32com.client #communication with COM ports
import pythoncom

VISITOR_HISTORY = r'C:\Users\Gabriel Rodenas\PycharmProjects\full_face_attendance_system\visitor_history'
temp_data_file = r"C:\Users\Gabriel Rodenas\PycharmProjects\full_face_attendance_system\temp_data.xlsx"


def remove_visitor_by_id(id):
    df = pd.read_csv(os.path.join(VISITOR_HISTORY, 'visitors_history.csv'))

    visitor_index = df[df['id'] == id].index

    if not visitor_index.empty:
        # Remove data from visitors_history.csv
        df = df.drop(visitor_index)
        df.to_csv(os.path.join(VISITOR_HISTORY, 'visitors_history.csv'), index=False)

        # Remove the visitor's picture from the directory
        visitor_picture = os.path.join(VISITOR_HISTORY, f'{id}.jpg')
        if os.path.isfile(visitor_picture):
            os.remove(visitor_picture)
            st.success(f'Visitor with id {id} removed successfully!')
        else:
            st.warning(f'No picture found for visitor with id {id}')
    else:
        st.error(f'No visitor found with id {id}')

#Disable Warnings
st.set_option('deprecation.showPyplotGlobalUse', False)
st.set_option('deprecation.showfileUploaderEncoding', False)


st.sidebar.markdown("""
                    > Made by Gabriel Rodenas & Rafael Rodenas (https://github.com/GabRod22)
                    """)

st.markdown('<h1><span style="color:gold;">HCPSMSHS</span> Face Attendance System</h1>', unsafe_allow_html=True)

#Defining Static Paths
if st.sidebar.button('Click to Clear out all the data'):
    #Clearing Visitor Database
    shutil.rmtree(VISITOR_DB, ignore_errors=True)
    os.mkdir(VISITOR_DB)
    #Clearing Visitor History
    shutil.rmtree(VISITOR_HISTORY, ignore_errors=True)
    os.mkdir(VISITOR_HISTORY)

if not os.path.exists(VISITOR_DB):
    os.mkdir(VISITOR_DB)

if not os.path.exists(VISITOR_HISTORY):
    os.mkdir(VISITOR_HISTORY)
# st.write(VISITOR_HISTORY)

def main():

    st.sidebar.header("About")
    st.sidebar.info("This web application is a Research project for Capstone and 3Is. ")


    selected_menu = option_menu(None,
        ['Visitor Validation', 'View Visitor History', 'Add to Database'],
        icons=['camera', "clock-history", 'person-plus'],
        ## icons from website: https://icons.getbootstrap.com/
        menu_icon="cast", default_index=0, orientation="horizontal")

    if selected_menu == 'Visitor Validation':
        ## Generates a Random ID for image storage
        visitor_id = uuid.uuid1()


        ## Reading Camera Image
        img_file_buffer = st.camera_input("Take a picture")

        if img_file_buffer is not None:
            bytes_data = img_file_buffer.getvalue()

            # convert image from opened file to np.array
            image_array         = cv2.imdecode(np.frombuffer(bytes_data,
                                                             np.uint8),
                                               cv2.IMREAD_COLOR)
            image_array_copy    = cv2.imdecode(np.frombuffer(bytes_data, np.uint8), cv2.IMREAD_COLOR)
            #st.image(cv2_img)

            #Saving Visitor History
            with open(os.path.join(VISITOR_HISTORY,
                                   f'{visitor_id}.jpg'), 'wb') as file:
                file.write(img_file_buffer.getbuffer())
                st.success('Image Saved Successfully!')

                #Validating Image
                # Detect faces in the loaded image
                # max_faces   = 0
                rois        = []  # region of interests (arrays of face areas)

                #To get location of Face from Image
                face_locations  = face_recognition.face_locations(image_array)
                #To encode Image to numeric format
                encodesCurFrame = face_recognition.face_encodings(image_array,
                                                                  face_locations)

                ## Generating Rectangle Red box over the Image
                for idx, (top, right, bottom, left) in enumerate(face_locations):
                    # Save face's Region of Interest
                    rois.append(image_array[top:bottom, left:right].copy())

                    # Draw a box around the face and label it
                    cv2.rectangle(image_array, (left, top), (right, bottom), COLOR_DARK, 2)
                    cv2.rectangle(image_array, (left, bottom + 35), (right, bottom), COLOR_DARK, cv2.FILLED)
                    font = cv2.FONT_HERSHEY_DUPLEX
                    cv2.putText(image_array, f"#{idx}", (left + 5, bottom + 25), font, .55, COLOR_WHITE, 1)

                # Showing Image
                # st.image(BGR_to_RGB(image_array), width=720)

                ## Number of Faces identified
                max_faces = len(face_locations)

                if max_faces > 0:

                    ## Set threshold for similarity
                    similarity_threshold = .90
                    flag_show = False

                    if (len(face_locations) > 0):
                        dataframe_new = pd.DataFrame()

                        for face_idx in range(max_faces):

                            ## Getting Region of Interest for that Face
                            roi = rois[face_idx]
                            # st.image(BGR_to_RGB(roi), width=min(roi.shape[0], 300))

                            # initial database for known faces
                            database_data = initialize_data()
                            # st.write(DB)

                            ## Getting Available information from Database
                            face_encodings  = database_data[COLS_ENCODE].values
                            dataframe       = database_data[COLS_INFO].copy()

                            # Comparing ROI to the faces available in database and finding distances and similarities
                            faces = face_recognition.face_encodings(roi)
                            # st.write(faces)

                            if len(faces) < 1:
                                ## Face could not be processed
                                st.error(f'Please Try Again for face#{face_idx}!')
                            else:
                                face_to_compare = faces[0]
                                ## Comparing Face with available information from database
                                dataframe['distance'] = face_recognition.face_distance(face_encodings,
                                                                                       face_to_compare)
                                dataframe['distance'] = dataframe['distance'].astype(float)

                                dataframe['similarity'] = dataframe.distance.apply(
                                    lambda distance: f"{face_distance_to_conf(distance):0.2}")
                                dataframe['similarity'] = dataframe['similarity'].astype(float)

                                dataframe_new = dataframe.drop_duplicates(keep='first')
                                dataframe_new.reset_index(drop=True, inplace=True)
                                dataframe_new.sort_values(by="similarity", ascending=True)

                                dataframe_new = dataframe_new[dataframe_new['similarity'] > similarity_threshold].head(1)
                                dataframe_new.reset_index(drop=True, inplace=True)

                                if dataframe_new.shape[0]>0:
                                    (top, right, bottom, left) = (face_locations[face_idx])

                                    ## Save Face Region of Interest information to the list
                                    rois.append(image_array_copy[top:bottom, left:right].copy())

                                    # Draw a Rectangle Red box around the face and label it
                                    cv2.rectangle(image_array_copy, (left, top), (right, bottom), COLOR_DARK, 2)
                                    cv2.rectangle(image_array_copy, (left, bottom + 35), (right, bottom), COLOR_DARK, cv2.FILLED)
                                    font = cv2.FONT_HERSHEY_DUPLEX
                                    cv2.putText(image_array_copy, f"#{dataframe_new.loc[0, 'Name']}", (left + 5, bottom + 25), font, .55, COLOR_WHITE, 1)

                                    ## Getting Name of Visitor
                                    name_visitor = dataframe_new.loc[0, 'Name']
                                    face_section = dataframe_new.loc[0, 'Section']
                                    attendance(visitor_id, name_visitor, face_section)

                                    flag_show = True

                                else:
                                    st.error(f'No Match Found for the given Similarity Threshold! for face#{face_idx}')
                                    st.info('Please Update the database for a new person or click again!')
                                    attendance(visitor_id, 'Unknown')

                        if flag_show:
                            st.image(BGR_to_RGB(image_array_copy), width=720)

                else:
                    st.error('No human face detected.')

            # Temperature Functions >:D
            # Specify the file path of the Excel file
            file_path = 'C:\\Users\\Gabriel Rodenas\\PycharmProjects\\full_face_attendance_system\\temp_data.xlsx'

            # Autosave interval in seconds
            autosave_interval = 5  # Change this value to your desired autosave interval

            # Initialize the COM library
            pythoncom.CoInitialize()

            # Create an instance of the Excel application
            excel = win32com.client.Dispatch("Excel.Application")

            # Open the Excel file
            workbook = excel.Workbooks.Open(file_path)

            # Create a Streamlit text element to display the cell value
            cell_value_text = st.empty()

            while True:
                # Save the Excel file
                workbook.Save()

                # Get the "Data In" sheet
                worksheet = workbook.Sheets("Data In")

                # Get the value of cell B22
                cell_value = worksheet.Range("B22").Value

                # Update the Streamlit text element with the cell value
                st.success(f"Temp Data: {cell_value}")

                # Wait for the specified interval before the next autosave
                time.sleep(autosave_interval)

    if selected_menu == 'View Visitor History':
        view_attendace()

        visitor_id = st.text_input("Enter the ID of the visitor to remove:")

        if st.button('Remove visitor'):
            remove_visitor_by_id(visitor_id)

    if selected_menu == 'Add to Database':
        col1, col2, col3 = st.columns(3)

        face_name  = col1.text_input('Name:', '')
        face_section = col2.text_input('Section:', '')
        pic_option = col2.radio('Upload Picture',
                                options=["Upload a Picture",
                                         "Click a picture"])

        if pic_option == 'Upload a Picture':
            img_file_buffer = col3.file_uploader('Upload a Picture',
                                                 type=allowed_image_type)
            if img_file_buffer is not None:
                # To read image file buffer with OpenCV:
                file_bytes = np.asarray(bytearray(img_file_buffer.read()),
                                        dtype=np.uint8)

        elif pic_option == 'Click a picture':
            img_file_buffer = col3.camera_input("Click a picture")
            if img_file_buffer is not None:
                # To read image file buffer with OpenCV:
                file_bytes = np.frombuffer(img_file_buffer.getvalue(),
                                           np.uint8)

        if ((img_file_buffer is not None) & (len(face_name) > 1) &
                st.button('Click to Save!')):
            # convert image from opened file to np.array
            image_array = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)
            # st.write(image_array)
            # st.image(cv2_img)

            with open(os.path.join(VISITOR_DB,
                                   f'{face_name}.jpg'), 'wb') as file:
                file.write(img_file_buffer.getbuffer())
                # st.success('Image Saved Successfully!')

            face_locations = face_recognition.face_locations(image_array)
            encodesCurFrame = face_recognition.face_encodings(image_array,
                                                              face_locations)

            df_new = pd.DataFrame(data=encodesCurFrame,
                                  columns=COLS_ENCODE)
            df_new[COLS_INFO] = [face_name, face_section]
            df_new = df_new[COLS_INFO + COLS_ENCODE].copy()

            # st.write(df_new)
            # initial database for known faces
            DB = initialize_data()
            add_data_db(df_new)


if __name__ == "__main__":
    main()
