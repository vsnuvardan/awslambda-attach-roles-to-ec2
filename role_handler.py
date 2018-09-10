import boto3
import xlrd
from botocore.exceptions import ClientError
from xlwt import Workbook
from custom_classes import Status
from custom_exceptions import RoleNotAttached, InstanceNotFound, InstanceTerminated
import  logging
from datetime import datetime

logger=logging.getLogger()

logger.setLevel(logging.INFO)


def role_handler(event,context):
    s3Client=boto3.client('s3')

    with open('/tmp/instanceids.xlsx','wb') as data:
        s3Client.download_fileobj('alexaccountlink','instanceids.xlsx',data)

    output=attch_role()
    jsonout=[]

    workbook = Workbook()
    worksheet = workbook.add_sheet("output")
    row_header=worksheet.row(0)
    row_header.write( 0, "INSTANCE ID")
    row_header.write(1, "REGION")
    row_header.write(2, "STATUS")
    row_marker=1
    for status in output:
        jsonout.append(status.toJSON())
        nrow=worksheet.row(row_marker)
        nrow.write(0,status.instanceId)
        nrow.write(1,status.region)
        nrow.write(2,status.status)
        row_marker=row_marker+1

    workbook.save('/tmp/output.xls')
    s3ObjectName='output'+str(datetime.now().date())+'-'+str(datetime.now().time()).replace(':','-').split('.')[0]+'.xls'
    with open('/tmp/output.xls', 'rb') as data:
        s3Client.upload_fileobj(data, 'alexaccountlink', s3ObjectName)



    return str(s3ObjectName+':'+str(jsonout))

def attch_role():
    workbook = xlrd.open_workbook('/tmp/instanceids.xlsx')
    worksheet = workbook.sheet_by_name('instances')

    listOfInstances = []
    noOfRows = len(worksheet.col_values(0))

    for i in range(1, noOfRows):
        try:

            instanceDetails = worksheet.row_values(i)
            instanceStatus = Status(instanceDetails[0],instanceDetails[1],"Pending")

            ec2Client=boto3.client('ec2',instanceDetails[1])

            describeInstance=ec2Client.describe_instances(InstanceIds=[instanceDetails[0].strip()])

            if len(describeInstance['Reservations']) == 0:

                raise InstanceNotFound
            else:
                print("INSTANCE DETAILS FOUND "+str(describeInstance))
                print("CHECK STATUS : "+str(describeInstance["Reservations"][0]["Instances"][0]['State']['Name']))
                if  describeInstance["Reservations"][0]["Instances"][0]['State']['Name'] == 'terminated':
                    raise InstanceTerminated({'error':'Instance Already Terminated'})


                if 'IamInstanceProfile' not in describeInstance["Reservations"][0]["Instances"][0].keys() :
                    print("====================" + str('IamInstanceProfile' in describeInstance["Reservations"][0]["Instances"][0].keys()))
                    raise RoleNotAttached({'message':'No Existing Role Found on the Instance'})
                
                else:

                    instnceProfile=describeInstance["Reservations"][0]["Instances"][0]["IamInstanceProfile"]
                    print("ROLE FOUND ON EC2 : "+str(instnceProfile))
                    iamResource=boto3.resource('iam')
                    readPolicy = iamResource.Role(instnceProfile['Arn'].split('/')[1]).attached_policies.all()
                    listOfPolicies=[]
                    for policy in list(readPolicy):
                        listOfPolicies.append(str(policy.arn).split("/")[1])
                    print("POLICIES : "+str(listOfPolicies))
                    print("CHECK IF COND : "+str('Platform' not in describeInstance["Reservations"][0]["Instances"][0].keys()  and 'cloudwatch-custom-Metrics-linux' in listOfPolicies))
                    if 'Platform' not in describeInstance["Reservations"][0]["Instances"][0].keys()  and 'cloudwatch-custom-Metrics-linux' in listOfPolicies:
                        message="ROLE ALREADY ATTACHED AND REQUIRED POLICY FOUND ON THE ROLE"

                    elif 'Platform' in describeInstance["Reservations"][0]["Instances"][0].keys() and 'cloudwatch-custom-Metrics-windows' in listOfPolicies:
                        message = "ROLE ALREADY ATTACHED AND REQUIRED POLICY FOUND ON THE ROLE"
                    else:
                        message="ROLE ALREADY ATTACHED BUT REQUIRED POLICY IS NOT FOUND ON THE ROLE"
        except InstanceTerminated as error:
            print("INSTANCE ALREADY GOT TERMINATED")
            message=str(error)
        except RoleNotAttached as error:
            print("NO ROLE FOUND ON EC2 ATTCHING THE ROLE")
            iamClient=boto3.client('iam')
            getRole=iamClient.get_role(RoleName='cloudwatch-exportcustommetrics')
            if 'Platform' not in describeInstance["Reservations"][0]["Instances"][0].keys() :
                associateInstanceProfile=ec2Client.associate_iam_instance_profile(IamInstanceProfile={
                    'Arn': 'arn:aws:iam::525064439504:instance-profile/cloudwatch-custom-Metrics-linux', 'Name': 'cloudwatch-custom-Metrics-linux'},InstanceId=instanceDetails[0].strip())
                message = "ROLE ATTACHED TO INSTANCE : cloudwatch-custom-Metrics-linux"
            else:
                associateInstanceProfile = ec2Client.associate_iam_instance_profile(IamInstanceProfile={
                    'Arn': 'arn:aws:iam::525064439504:instance-profile/cloudwatch-custom-Metrics-windows',
                    'Name': 'cloudwatch-custom-Metrics-windows'}, InstanceId=instanceDetails[0].strip())
                message = "ROLE ATTACHED TO INSTANCE : cloudwatch-custom-Metrics-windows"


        except KeyError as error:
            print("EXCEPTION OCCURED KEY ERROR : "+str(error))

        except IndexError as error:
            print("INDEX ERROR  OCCURED : "+str(error))
            message=str(error)
        except InstanceNotFound as error:
            print("INSTANCE NOT FOUND ")
            message="NO INSTANCE FOUND WITH THE GIVEN ID "

        except ClientError as error:
            print("CLIENT ERROR OCCURED : "+str(error))
            message=str(error)

        except Exception as error:
            message=str(error)
            print("EXCEPTION OCCURED : "+str(error))

        print("OUTPUT MESSAGE : "+message)

        instanceStatus.status=message
        listOfInstances.append(instanceStatus)

    return listOfInstances



print("==="+str(role_handler("","")))