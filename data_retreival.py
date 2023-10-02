from flask import Flask, jsonify, request
import pandas as pd
app = Flask(__name__)

excel_path = r'C:\Users\Ojal\Documents\IT asset mngr\main\new\ack.xlsx'
ack_data = pd.DataFrame(columns=['IP Adress','Ack Status'])

@app.route('/acknowledge', methods=['GET'])
def acknowledge():
    acknowledgment_status = request.args.get('status')

    # Process the acknowledgment_status (either "yes" or "no")
    if acknowledgment_status not in ['yes', 'no']:
        return "Invalid acknowledgment status"
    
    ip_address = request.remote_addr
    ack_data.loc[len(ack_data)] = [ip_address,acknowledgment_status] #len(ack_data) returns the number of rows

    ack_data.to_excel(excel_path,index=False)

    return f"\t Thankyou, Acknoledgement Received:{acknowledgment_status}"
def get_ack():
    ip_address = request.remote_addr
    ack_status = ack_data.loc[ack_data['IP Address'] == ip_address, 'Acknowledgment Status'].values
    ack_status = ack_status[0] if len(ack_status)>0 else 'Not Acknowleged'

    return jsonify({"acknowledgement_status":ack_status})

if __name__ == '__main__':
    try:
        ack_data = pd.read_excel(excel_path)
    except FileNotFoundError:
        pass

    app.run(host='0.0.0.0', port=80)  # Replace host and port with your preferences.
