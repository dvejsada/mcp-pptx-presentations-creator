import boto3
from botocore.exceptions import NoCredentialsError, ClientError
import uuid
import os

def generate_unique_object_name():
    """Generate a unique object name using UUID and preserve the file extension.

    :return: Unique object name with the same file extension
    """

    # Generate a UUID
    unique_id = str(uuid.uuid4())
    # Combine UUID and extension
    unique_object_name = f"{unique_id}.pptx"

    return unique_object_name

def upload_file_to_s3(file_name):
    """Upload a file to an S3 bucket and return a pre-signed URL valid for 1 hour.

    :param file_name: Path to the file to upload
    :param bucket: Name of the S3 bucket
    :param object_name: S3 object name (optional). If not specified, file_name is used
    :param region: AWS region where the bucket is hosted
    :return: Pre-signed URL string if successful, else None
    """
    aws_access_key_id = os.environ["aws_access_key_id"]
    aws_secret_access_key = os.environ["aws_secret_access_key"]
    aws_region= os.environ['aws_region']
    aws_bucket = os.environ["aws_bucket"]

    object_name = generate_unique_object_name()

    # Create an S3 client
    s3_client = boto3.client('s3', region_name=aws_region, aws_access_key_id= aws_access_key_id,
    aws_secret_access_key= aws_secret_access_key, endpoint_url='https://s3.eu-central-1.amazonaws.com')

    try:
        # Upload the file to S3
        s3_client.upload_file(file_name, aws_bucket, object_name)

        # Generate a pre-signed URL valid for 1 hour (3600 seconds)
        url = s3_client.generate_presigned_url('get_object',
                                               Params={'Bucket': aws_bucket,
                                                       'Key': object_name},
                                               ExpiresIn=3600)
        return url

    except FileNotFoundError:
        print(f"The file {file_name} was not found.")
        return None
    except NoCredentialsError:
        print("AWS credentials are not available.")
        return None
    except ClientError as e:
        print(f"Client error: {e}")
        return None
