import os
import sys
import threading
import webbrowser
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import pandas as pd
from datetime import datetime
import logging
from pathlib import Path
import socket
import traceback
import subprocess

# ===== DEFAULT CREDENTIALS CONFIGURATION =====
# Set your default Windows credentials here
DEFAULT_USERNAME = "trf"  # e.g., "DOMAIN\\admin" or "admin"
DEFAULT_PASSWORD = "trf"  # Your network password
USE_DEFAULT_CREDENTIALS = True  # Set to False to disable auto-authentication
# =============================================

# Determine if we're running as a PyInstaller bundle
if getattr(sys, 'frozen', False):
    # Running as compiled executable
    application_path = sys._MEIPASS
    base_path = os.path.dirname(sys.executable)
    print(f"Running as EXE - Application path: {application_path}")
    print(f"Running as EXE - Base path: {base_path}")
else:
    # Running as script
    application_path = os.path.dirname(os.path.abspath(__file__))
    base_path = application_path
    print(f"Running as script - Application path: {application_path}")
    print(f"Running as script - Base path: {base_path}")

# Set up Flask with correct paths
template_folder = os.path.join(application_path, 'templates')
static_folder = os.path.join(application_path, 'assets')
print(f"Template folder: {template_folder}")
print(f"Template folder exists: {os.path.exists(template_folder)}")
print(f"Static folder: {static_folder}")

if os.path.exists(template_folder):
    print(f"Contents of template folder: {os.listdir(template_folder)}")

app = Flask(__name__, template_folder=template_folder, static_folder=static_folder, static_url_path='/assets')

# Create upload and report folders in the executable's directory
app.config['UPLOAD_FOLDER'] = os.path.join(base_path, 'uploads')
app.config['REPORT_FOLDER'] = os.path.join(base_path, 'reports')
app.config['DATA_FOLDER'] = os.path.join(base_path, 'data')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Create necessary folders
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['REPORT_FOLDER'], exist_ok=True)
os.makedirs(app.config['DATA_FOLDER'], exist_ok=True)

print(f"Upload folder: {app.config['UPLOAD_FOLDER']}")
print(f"Report folder: {app.config['REPORT_FOLDER']}")

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Store active network connections
active_connections = {}


def get_credentials(username=None, password=None):
    """
    Get credentials - use provided ones or fall back to defaults
    Returns: tuple (username, password)
    """
    if username and password:
        return username, password
    elif USE_DEFAULT_CREDENTIALS:
        logger.info("Using default credentials")
        return DEFAULT_USERNAME, DEFAULT_PASSWORD
    else:
        return None, None


def connect_to_network_share(ip_address, username=None, password=None):
    """
    Establish a network connection using net use command
    Uses default credentials if none provided and USE_DEFAULT_CREDENTIALS is True
    Returns: dict with success status and message
    """
    try:
        # Get credentials (use defaults if not provided)
        username, password = get_credentials(username, password)
        
        # Construct the network path
        if ip_address.startswith('\\\\'):
            network_path = ip_address
        else:
            network_path = f'\\\\{ip_address}'
        
        # Check if already connected
        connection_key = f"{ip_address}_{username}" if username else ip_address
        if connection_key in active_connections:
            logger.info(f"Already connected to {network_path}")
            return {'success': True, 'message': 'Déjà connecté'}
        
        # Disconnect first if there's an existing connection (without credentials)
        disconnect_cmd = f'net use {network_path} /delete /y'
        subprocess.run(disconnect_cmd, shell=True, capture_output=True, text=True)
        
        # Build the net use command
        if username and password:
            connect_cmd = f'net use {network_path} /user:{username} {password}'
            logger.info(f"Connecting to {network_path} with user: {username}")
        else:
            connect_cmd = f'net use {network_path}'
            logger.info(f"Connecting to {network_path} without credentials")
        
        # Execute the command
        result = subprocess.run(
            connect_cmd, 
            shell=True, 
            capture_output=True, 
            text=True,
            timeout=10
        )
        
        if result.returncode == 0:
            active_connections[connection_key] = {
                'ip': ip_address,
                'username': username or 'default',
                'connected_at': datetime.now()
            }
            logger.info(f"Successfully connected to {network_path}")
            return {'success': True, 'message': 'Connexion réussie'}
        else:
            error_msg = result.stderr.strip() or result.stdout.strip()
            logger.error(f"Failed to connect to {network_path}: {error_msg}")
            return {'success': False, 'message': f'Échec de connexion: {error_msg}'}
    
    except subprocess.TimeoutExpired:
        logger.error(f"Connection timeout for {ip_address}")
        return {'success': False, 'message': 'Délai de connexion dépassé'}
    except Exception as e:
        logger.error(f"Error connecting to {ip_address}: {str(e)}")
        return {'success': False, 'message': str(e)}


def disconnect_from_network_share(ip_address):
    """
    Disconnect from a network share
    """
    try:
        if ip_address.startswith('\\\\'):
            network_path = ip_address
        else:
            network_path = f'\\\\{ip_address}'
        
        disconnect_cmd = f'net use {network_path} /delete /y'
        subprocess.run(disconnect_cmd, shell=True, capture_output=True, text=True)
        
        # Remove from active connections
        keys_to_remove = [k for k in active_connections if active_connections[k]['ip'] == ip_address]
        for key in keys_to_remove:
            del active_connections[key]
        
        logger.info(f"Disconnected from {network_path}")
        return True
    except Exception as e:
        logger.error(f"Error disconnecting from {ip_address}: {str(e)}")
        return False


def check_file_exists(ip_address, directory_path, filename, username=None, password=None):
    """
    Check if a file exists on a network path with authentication
    Uses default credentials if none provided
    Returns: dict with status and details
    """
    try:
        # Get credentials (use defaults if not provided)
        username, password = get_credentials(username, password)
        
        # Connect to the network share first if credentials available
        if username and password:
            connection_result = connect_to_network_share(ip_address, username, password)
            
            if not connection_result['success']:
                return {
                    'exists': False,
                    'path': f'\\\\{ip_address}\\{directory_path}\\{filename}',
                    'size': None,
                    'modified': None,
                    'error': f"Échec de connexion: {connection_result['message']}"
                }
        
        # Construct the network path
        if ip_address.startswith('\\\\'):
            network_path = os.path.join(ip_address, directory_path, filename)
        else:
            network_path = os.path.join(f'\\\\{ip_address}', directory_path, filename)
        
        # Normalize the path
        network_path = os.path.normpath(network_path)
        
        # Check if file exists
        exists = os.path.exists(network_path)
        
        if exists:
            # Get file size
            size = os.path.getsize(network_path)
            modified_time = datetime.fromtimestamp(os.path.getmtime(network_path))
            
            return {
                'exists': True,
                'path': network_path,
                'size': size,
                'modified': modified_time.strftime('%Y-%m-%d %H:%M:%S'),
                'error': None
            }
        else:
            return {
                'exists': False,
                'path': network_path,
                'size': None,
                'modified': None,
                'error': 'File not found'
            }
    except Exception as e:
        logger.error(f"Error checking {ip_address}/{directory_path}/{filename}: {str(e)}")
        return {
            'exists': False,
            'path': f'\\\\{ip_address}\\{directory_path}\\{filename}',
            'size': None,
            'modified': None,
            'error': str(e)
        }


def process_excel(file_path, filename_to_check, directory_path, username=None, password=None):
    """
    Process the uploaded Excel file and check for file existence
    Uses default credentials if none provided
    """
    try:
        # Read Excel file
        df = pd.read_excel(file_path)
        
        # Validate required columns
        if 'CodeMag' not in df.columns or 'ipaddress' not in df.columns:
            return {'error': 'Excel must contain "CodeMag" and "ipaddress" columns'}
        
        results = []
        total = len(df)
        
        for index, row in df.iterrows():
            code_mag = str(row['CodeMag'])
            ip_address = str(row['ipaddress'])
            
            # Check file existence (will use default credentials if none provided)
            result = check_file_exists(ip_address, directory_path, filename_to_check, username, password)
            
            results.append({
                'CodeMag': code_mag,
                'IPAddress': ip_address,
                'FileName': filename_to_check,
                'Exists': 'Yes' if result['exists'] else 'No',
                'FilePath': result['path'],
                'FileSize': result['size'],
                'LastModified': result['modified'],
                'Error': result['error']
            })
            
            # Log progress
            logger.info(f"Processed {index + 1}/{total}: {code_mag} - {ip_address}")
        
        return {'success': True, 'results': results}
        
    except Exception as e:
        logger.error(f"Error processing Excel: {str(e)}")
        return {'error': str(e)}


def get_excel_files_from_data():
    """
    Get list of Excel files from the data folder
    """
    try:
        data_folder = app.config['DATA_FOLDER']
        excel_files = []
        
        if os.path.exists(data_folder):
            for file in os.listdir(data_folder):
                if file.endswith(('.xlsx', '.xls')):
                    excel_files.append(file)
        
        return excel_files
    except Exception as e:
        logger.error(f"Error getting Excel files: {str(e)}")
        return []


def transfer_file_to_servers(file_path, servers_excel, directory_path, username=None, password=None):
    """
    Transfer a file to multiple servers based on Excel file
    Uses default credentials if none provided
    """
    try:
        # Get credentials (use defaults if not provided)
        username, password = get_credentials(username, password)
        
        # Read Excel file
        df = pd.read_excel(servers_excel)
        
        # Validate required columns
        if 'CodeMag' not in df.columns or 'ipaddress' not in df.columns:
            return {'error': 'Excel must contain "CodeMag" and "ipaddress" columns'}
        
        results = []
        total = len(df)
        filename = os.path.basename(file_path)
        
        for index, row in df.iterrows():
            code_mag = str(row['CodeMag'])
            ip_address = str(row['ipaddress'])
            
            # Connect to the network share first if credentials available
            if username and password:
                connection_result = connect_to_network_share(ip_address, username, password)
                
                if not connection_result['success']:
                    results.append({
                        'CodeMag': code_mag,
                        'IPAddress': ip_address,
                        'FileName': filename,
                        'Status': 'Failed',
                        'DestinationPath': f'\\\\{ip_address}\\{directory_path}\\{filename}',
                        'Error': f"Échec de connexion: {connection_result['message']}"
                    })
                    continue
            
            # Construct destination path
            if ip_address.startswith('\\\\'):
                dest_path = os.path.join(ip_address, directory_path, filename)
            else:
                dest_path = os.path.join(f'\\\\{ip_address}', directory_path, filename)
            
            dest_path = os.path.normpath(dest_path)
            
            try:
                # Create directory if it doesn't exist
                dest_dir = os.path.dirname(dest_path)
                os.makedirs(dest_dir, exist_ok=True)
                
                # Copy file
                import shutil
                shutil.copy2(file_path, dest_path)
                
                results.append({
                    'CodeMag': code_mag,
                    'IPAddress': ip_address,
                    'FileName': filename,
                    'Status': 'Success',
                    'DestinationPath': dest_path,
                    'Error': None
                })
                
                logger.info(f"Transferred {index + 1}/{total}: {code_mag} - {ip_address}")
                
            except Exception as e:
                logger.error(f"Error transferring to {ip_address}: {str(e)}")
                results.append({
                    'CodeMag': code_mag,
                    'IPAddress': ip_address,
                    'FileName': filename,
                    'Status': 'Failed',
                    'DestinationPath': dest_path,
                    'Error': str(e)
                })
        
        return {'success': True, 'results': results}
        
    except Exception as e:
        logger.error(f"Error in file transfer: {str(e)}")
        return {'error': str(e)}


@app.route('/')
def index():
    try:
        logger.info("Accessing index route")
        logger.info(f"Looking for template in: {app.template_folder}")
        return render_template('index.html')
    except Exception as e:
        logger.error(f"Error in index route: {str(e)}")
        logger.error(traceback.format_exc())
        return f"""
        <html>
        <body>
        <h1>Error Loading Application</h1>
        <p>Error: {str(e)}</p>
        <pre>{traceback.format_exc()}</pre>
        <p>Template folder: {app.template_folder}</p>
        <p>Template folder exists: {os.path.exists(app.template_folder)}</p>
        </body>
        </html>
        """, 500


@app.errorhandler(500)
def internal_error(error):
    logger.error(f"500 Error: {str(error)}")
    logger.error(traceback.format_exc())
    return f"""
    <html>
    <body>
    <h1>Internal Server Error</h1>
    <p>Error: {str(error)}</p>
    <pre>{traceback.format_exc()}</pre>
    </body>
    </html>
    """, 500


@app.errorhandler(Exception)
def handle_exception(e):
    logger.error(f"Unhandled exception: {str(e)}")
    logger.error(traceback.format_exc())
    return f"""
    <html>
    <body>
    <h1>Unhandled Exception</h1>
    <p>Error: {str(e)}</p>
    <pre>{traceback.format_exc()}</pre>
    </body>
    </html>
    """, 500


@app.route('/upload', methods=['POST'])
def upload_file():
    """Legacy route - kept for backwards compatibility"""
    try:
        # Check if file was uploaded
        if 'excel_file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['excel_file']
        filename_to_check = request.form.get('filename')
        directory_path = request.form.get('directory_path')
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not filename_to_check:
            return jsonify({'error': 'Please specify the filename to check'}), 400
        
        if not directory_path:
            return jsonify({'error': 'Please specify the directory path'}), 400
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Process the Excel file (will use default credentials)
        result = process_excel(filepath, filename_to_check, directory_path)
        
        if 'error' in result:
            return jsonify({'error': result['error']}), 400
        
        # Generate report
        report_filename = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        report_path = os.path.join(app.config['REPORT_FOLDER'], report_filename)
        
        # Create DataFrame from results
        df_report = pd.DataFrame(result['results'])
        
        # Save to Excel with formatting
        with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
            df_report.to_excel(writer, sheet_name='Results', index=False)
            
            # Get workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Results']
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
        
        # Calculate summary
        total_checked = len(result['results'])
        found = sum(1 for r in result['results'] if r['Exists'] == 'Yes')
        not_found = total_checked - found
        
        return jsonify({
            'success': True,
            'report_file': report_filename,
            'summary': {
                'total': total_checked,
                'found': found,
                'not_found': not_found
            },
            'results': result['results']
        })
        
    except Exception as e:
        logger.error(f"Upload error: {str(e)}")
        return jsonify({'error': str(e)}), 500


@app.route('/check-files', methods=['POST'])
def check_files():
    """New route that uses Excel files from data folder"""
    try:
        excel_filename = request.form.get('excel_file')
        filename_to_check = request.form.get('filename')
        directory_path = request.form.get('directory_path')
        
        if not excel_filename:
            return jsonify({'error': 'Veuillez sélectionner un fichier Excel'}), 400
        
        if not filename_to_check:
            return jsonify({'error': 'Veuillez spécifier le nom du fichier à vérifier'}), 400
        
        if not directory_path:
            return jsonify({'error': 'Veuillez spécifier le chemin du répertoire'}), 400
        
        # Get Excel file path from data folder
        excel_path = os.path.join(app.config['DATA_FOLDER'], excel_filename)
        
        if not os.path.exists(excel_path):
            return jsonify({'error': f'Fichier Excel non trouvé: {excel_filename}'}), 400
        
        # Process the Excel file (will use default credentials)
        result = process_excel(excel_path, filename_to_check, directory_path)
        
        if 'error' in result:
            return jsonify({'error': result['error']}), 400
        
        # Generate report
        report_filename = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        report_path = os.path.join(app.config['REPORT_FOLDER'], report_filename)
        
        # Create DataFrame from results
        df_report = pd.DataFrame(result['results'])
        
        # Save to Excel with formatting
        with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
            df_report.to_excel(writer, sheet_name='Results', index=False)
            
            # Get workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Results']
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
        
        # Calculate summary
        total_checked = len(result['results'])
        found = sum(1 for r in result['results'] if r['Exists'] == 'Yes')
        not_found = total_checked - found
        
        return jsonify({
            'success': True,
            'report_file': report_filename,
            'summary': {
                'total': total_checked,
                'found': found,
                'not_found': not_found
            },
            'results': result['results']
        })
        
    except Exception as e:
        logger.error(f"Check files error: {str(e)}")
        return jsonify({'error': str(e)}), 500


@app.route('/download/<filename>')
def download_report(filename):
    try:
        report_path = os.path.join(app.config['REPORT_FOLDER'], filename)
        return send_file(report_path, as_attachment=True)
    except Exception as e:
        logger.error(f"Download error: {str(e)}")
        return jsonify({'error': str(e)}), 404

@app.route('/transfer-files', methods=['POST'])
def transfer_files():
    """Transfer multiple files to servers"""
    try:
        # Check if files were uploaded
        if 'files_to_transfer' not in request.files:
            return jsonify({'error': 'No files uploaded'}), 400
        
        files = request.files.getlist('files_to_transfer')
        excel_filename = request.form.get('excel_file')
        directory_path = request.form.get('directory_path')
        
        # Validate files
        valid_files = []
        for file in files:
            if file.filename != '':
                valid_files.append(file)
        
        if len(valid_files) == 0:
            return jsonify({'error': 'No files selected'}), 400
        
        if not excel_filename:
            return jsonify({'error': 'Please select an Excel file'}), 400
        
        if not directory_path:
            return jsonify({'error': 'Please specify the directory path'}), 400
        
        # Get Excel file path from data folder
        excel_path = os.path.join(app.config['DATA_FOLDER'], excel_filename)
        
        if not os.path.exists(excel_path):
            return jsonify({'error': f'Excel file not found: {excel_filename}'}), 400
        
        # Save uploaded files temporarily
        temp_filepaths = []
        for file in valid_files:
            filename = secure_filename(file.filename)
            temp_filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(temp_filepath)
            temp_filepaths.append(temp_filepath)
        
        # Transfer all files (will use default credentials)
        all_results = []
        for temp_filepath in temp_filepaths:
            result = transfer_files_to_servers(temp_filepath, excel_path, directory_path)
            if 'error' not in result:
                all_results.extend(result['results'])
            else:
                # If one file fails, log it but continue with others
                logger.error(f"Error transferring file {temp_filepath}: {result['error']}")
        
        # Generate report
        report_filename = f"transfer_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        report_path = os.path.join(app.config['REPORT_FOLDER'], report_filename)
        
        # Create DataFrame from all results
        df_report = pd.DataFrame(all_results)
        
        # Save to Excel with formatting
        with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
            df_report.to_excel(writer, sheet_name='Transfer Results', index=False)
            
            # Get workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Transfer Results']
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
        
        # Calculate summary
        total_transfers = len(all_results)
        successful = sum(1 for r in all_results if r['Status'] == 'Success')
        failed = total_transfers - successful
        
        # Clean up temporary files
        for temp_filepath in temp_filepaths:
            try:
                os.remove(temp_filepath)
            except:
                pass
        
        return jsonify({
            'success': True,
            'report_file': report_filename,
            'summary': {
                'total': total_transfers,
                'successful': successful,
                'failed': failed
            },
            'results': all_results
        })
        
    except Exception as e:
        logger.error(f"Transfer files error: {str(e)}")
        return jsonify({'error': str(e)}), 500


def transfer_files_to_servers(file_path, servers_excel, directory_path, username=None, password=None):
    """
    Transfer a file to multiple servers based on Excel file
    Uses default credentials if none provided
    """
    try:
        # Get credentials (use defaults if not provided)
        username, password = get_credentials(username, password)
        
        # Read Excel file
        df = pd.read_excel(servers_excel)
        
        # Validate required columns
        if 'CodeMag' not in df.columns or 'ipaddress' not in df.columns:
            return {'error': 'Excel must contain "CodeMag" and "ipaddress" columns'}
        
        results = []
        total = len(df)
        filename = os.path.basename(file_path)
        
        for index, row in df.iterrows():
            code_mag = str(row['CodeMag'])
            ip_address = str(row['ipaddress'])
            
            # Connect to the network share first if credentials available
            if username and password:
                connection_result = connect_to_network_share(ip_address, username, password)
                
                if not connection_result['success']:
                    results.append({
                        'CodeMag': code_mag,
                        'IPAddress': ip_address,
                        'FileName': filename,
                        'Status': 'Failed',
                        'DestinationPath': f'\\\\{ip_address}\\{directory_path}\\{filename}',
                        'Error': f"Échec de connexion: {connection_result['message']}"
                    })
                    continue
            
            # Construct destination path
            if ip_address.startswith('\\\\'):
                dest_path = os.path.join(ip_address, directory_path, filename)
            else:
                dest_path = os.path.join(f'\\\\{ip_address}', directory_path, filename)
            
            dest_path = os.path.normpath(dest_path)
            
            try:
                # Create directory if it doesn't exist
                dest_dir = os.path.dirname(dest_path)
                os.makedirs(dest_dir, exist_ok=True)
                
                # Copy file
                import shutil
                shutil.copy2(file_path, dest_path)
                
                results.append({
                    'CodeMag': code_mag,
                    'IPAddress': ip_address,
                    'FileName': filename,
                    'Status': 'Success',
                    'DestinationPath': dest_path,
                    'Error': None
                })
                
                logger.info(f"Transferred {index + 1}/{total}: {code_mag} - {ip_address} - {filename}")
                
            except Exception as e:
                logger.error(f"Error transferring {filename} to {ip_address}: {str(e)}")
                results.append({
                    'CodeMag': code_mag,
                    'IPAddress': ip_address,
                    'FileName': filename,
                    'Status': 'Failed',
                    'DestinationPath': dest_path,
                    'Error': str(e)
                })
        
        return {'success': True, 'results': results}
        
    except Exception as e:
        logger.error(f"Error in file transfer: {str(e)}")
        return {'error': str(e)}
    
@app.route('/auth')
def auth_page():
    try:
        logger.info("Accessing auth page")
        return render_template('auth.html')
    except Exception as e:
        logger.error(f"Error in auth route: {str(e)}")
        logger.error(traceback.format_exc())
        return f"""
        <html>
        <body>
        <h1>Error Loading Authentication Page</h1>
        <p>Error: {str(e)}</p>
        <pre>{traceback.format_exc()}</pre>
        </body>
        </html>
        """, 500


@app.route('/test-connection', methods=['POST'])
def test_connection():
    """Test network connection with credentials"""
    try:
        data = request.get_json()
        ip_address = data.get('ip_address')
        username = data.get('username')
        password = data.get('password')
        
        if not ip_address:
            return jsonify({'error': 'Adresse IP requise'}), 400
        
        # Will use default credentials if username/password not provided
        result = connect_to_network_share(ip_address, username, password)
        
        if result['success']:
            cred_info = " (credentials par défaut)" if (not username and USE_DEFAULT_CREDENTIALS) else ""
            return jsonify({
                'success': True,
                'message': f'Connexion réussie à \\\\{ip_address}{cred_info}'
            })
        else:
            return jsonify({
                'success': False,
                'message': result['message']
            }), 400
    
    except Exception as e:
        logger.error(f"Test connection error: {str(e)}")
        return jsonify({'error': str(e)}), 500


@app.route('/disconnect-all', methods=['POST'])
def disconnect_all():
    """Disconnect all network shares"""
    try:
        count = len(active_connections)
        for connection_key in list(active_connections.keys()):
            ip = active_connections[connection_key]['ip']
            disconnect_from_network_share(ip)
        
        return jsonify({
            'success': True,
            'message': f'{count} connexion(s) fermée(s)'
        })
    
    except Exception as e:
        logger.error(f"Disconnect error: {str(e)}")
        return jsonify({'error': str(e)}), 500


@app.route('/test-bulk-connections', methods=['POST'])
def test_bulk_connections():
    """Test multiple network connections from Excel file"""
    try:
        data = request.get_json()
        excel_filename = data.get('excel_file', '')
        username = data.get('username', '').strip()
        password = data.get('password', '')
        
        if not excel_filename:
            return jsonify({'error': 'Fichier Excel requis'}), 400
        
        if not username or not password:
            return jsonify({'error': 'Nom d\'utilisateur et mot de passe requis'}), 400
        
        # Get Excel file path from data folder
        excel_path = os.path.join(app.config['DATA_FOLDER'], excel_filename)
        
        if not os.path.exists(excel_path):
            return jsonify({'error': f'Fichier Excel introuvable: {excel_filename}'}), 400
        
        # Read Excel file
        logger.info(f"Reading Excel file: {excel_path}")
        df = pd.read_excel(excel_path)
        
        # Find the IP address column (flexible matching)
        ip_column = None
        for col in df.columns:
            col_lower = str(col).lower()
            if ('ip' in col_lower and 'address' in col_lower) or col_lower == 'ip' or col_lower == 'ip address' or col_lower == 'adresse ip':
                ip_column = col
                break
        
        if ip_column is None:
            # List available columns for debugging
            available_columns = ', '.join([str(col) for col in df.columns])
            return jsonify({'error': f'Aucune colonne "IP Address" trouvée. Colonnes disponibles: {available_columns}'}), 400
        
        # Get unique IP addresses
        ip_addresses = df[ip_column].dropna().unique().tolist()
        
        if len(ip_addresses) == 0:
            return jsonify({'error': 'Aucune adresse IP trouvée dans le fichier Excel'}), 400
        
        logger.info(f"Found {len(ip_addresses)} IP addresses to test")
        
        # Test each connection
        results = []
        successful = 0
        failed = 0
        
        for ip_address in ip_addresses:
            ip_str = str(ip_address).strip()
            if not ip_str or ip_str.lower() in ['nan', 'none', '']:
                continue
                
            logger.info(f"Testing connection to {ip_str}")
            result = connect_to_network_share(ip_str, username, password)
            
            if result['success']:
                successful += 1
                status = 'Success'
                message = result['message']
            else:
                failed += 1
                status = 'Failed'
                message = result['message']
            
            results.append({
                'ip_address': ip_str,
                'status': status,
                'message': message
            })
        
        logger.info(f"Bulk test completed: {successful} successful, {failed} failed out of {len(results)} total")
        
        return jsonify({
            'success': True,
            'summary': {
                'total': len(results),
                'successful': successful,
                'failed': failed
            },
            'results': results
        })
    
    except Exception as e:
        logger.error(f"Bulk test error: {str(e)}")
        logger.error(traceback.format_exc())
        return jsonify({'error': str(e)}), 500


@app.route('/active-connections')
def get_active_connections():
    """Get list of active connections"""
    try:
        connections = []
        for key, conn in active_connections.items():
            connections.append({
                'ip': conn['ip'],
                'username': conn['username'],
                'connected_at': conn['connected_at'].strftime('%Y-%m-%d %H:%M:%S')
            })
        
        return jsonify({
            'success': True,
            'connections': connections
        })
    
    except Exception as e:
        logger.error(f"Get connections error: {str(e)}")
        return jsonify({'error': str(e)}), 500


@app.route('/transfer')
def transfer_page():
    try:
        logger.info("Accessing transfer page")
        return render_template('transfer.html')
    except Exception as e:
        logger.error(f"Error in transfer route: {str(e)}")
        logger.error(traceback.format_exc())
        return f"""
        <html>
        <body>
        <h1>Error Loading Transfer Page</h1>
        <p>Error: {str(e)}</p>
        <pre>{traceback.format_exc()}</pre>
        </body>
        </html>
        """, 500


@app.route('/get-excel-files')
def get_excel_files():
    try:
        files = get_excel_files_from_data()
        return jsonify({'success': True, 'files': files})
    except Exception as e:
        logger.error(f"Error getting Excel files: {str(e)}")
        return jsonify({'error': str(e)}), 500


@app.route('/transfer-file', methods=['POST'])
def transfer_file():
    try:
        # Check if file was uploaded
        if 'file_to_transfer' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file_to_transfer']
        excel_filename = request.form.get('excel_file')
        directory_path = request.form.get('directory_path')
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not excel_filename:
            return jsonify({'error': 'Please select an Excel file'}), 400
        
        if not directory_path:
            return jsonify({'error': 'Please specify the directory path'}), 400
        
        # Save uploaded file temporarily
        filename = secure_filename(file.filename)
        temp_filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(temp_filepath)
        
        # Get Excel file path from data folder
        excel_path = os.path.join(app.config['DATA_FOLDER'], excel_filename)
        
        if not os.path.exists(excel_path):
            return jsonify({'error': f'Excel file not found: {excel_filename}'}), 400
        
        # Transfer the file (will use default credentials)
        result = transfer_file_to_servers(temp_filepath, excel_path, directory_path)
        
        if 'error' in result:
            return jsonify({'error': result['error']}), 400
        
        # Generate report
        report_filename = f"transfer_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        report_path = os.path.join(app.config['REPORT_FOLDER'], report_filename)
        
        # Create DataFrame from results
        df_report = pd.DataFrame(result['results'])
        
        # Save to Excel with formatting
        with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
            df_report.to_excel(writer, sheet_name='Transfer Results', index=False)
            
            # Get workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Transfer Results']
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
        
        # Calculate summary
        total_transfers = len(result['results'])
        successful = sum(1 for r in result['results'] if r['Status'] == 'Success')
        failed = total_transfers - successful
        
        # Clean up temporary file
        try:
            os.remove(temp_filepath)
        except:
            pass
        
        return jsonify({
            'success': True,
            'report_file': report_filename,
            'summary': {
                'total': total_transfers,
                'successful': successful,
                'failed': failed
            },
            'results': result['results']
        })
        
    except Exception as e:
        logger.error(f"Transfer error: {str(e)}")
        return jsonify({'error': str(e)}), 500


def open_browser():
    """Open the browser after a short delay"""
    import time
    time.sleep(1.5)
    webbrowser.open('http://127.0.0.1:5001')


if __name__ == '__main__':
    try:
        print("=" * 60)
        print("File Checker Application")
        print("=" * 60)
        print(f"Python version: {sys.version}")
        print(f"Flask app: {app}")
        print(f"Template folder: {app.template_folder}")
        print(f"Upload folder: {app.config['UPLOAD_FOLDER']}")
        print(f"Report folder: {app.config['REPORT_FOLDER']}")
        print(f"Default credentials enabled: {USE_DEFAULT_CREDENTIALS}")
        if USE_DEFAULT_CREDENTIALS:
            print(f"Default username: {DEFAULT_USERNAME}")
        print("=" * 60)
        
        # Open browser in a separate thread
        threading.Thread(target=open_browser, daemon=True).start()
        
        # Run Flask app on port 5001 to avoid conflicts
        print("Starting Flask server on http://127.0.0.1:5001")
        print("Press CTRL+C to stop the server")
        print("=" * 60)
        app.run(host='127.0.0.1', port=5001, debug=False, use_reloader=False)
    except Exception as e:
        print(f"\n{'='*60}")
        print("CRITICAL ERROR:")
        print(f"{'='*60}")
        print(str(e))
        print(traceback.format_exc())
        print(f"{'='*60}")
        input("Press Enter to exit...")