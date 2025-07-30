import boto3
import json
import re
from datetime import datetime, timedelta
from datetime import datetime, date
from io import BytesIO
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from botocore.exceptions import ClientError
from decimal import Decimal
import base64
import traceback
import os
import uuid

# Environment variables for configuration
BEDROCK_MODEL_ID = os.environ.get('BEDROCK_MODEL_ID', 'anthropic.claude-3-sonnet-20240229-v1:0')
BEDROCK_TEMPERATURE = float(os.environ.get('BEDROCK_TEMPERATURE', '0.1'))
BEDROCK_TOP_P = float(os.environ.get('BEDROCK_TOP_P', '0.9'))
BEDROCK_MAX_TOKENS = int(os.environ.get('BEDROCK_MAX_TOKENS', '4000'))
EXCEL_FILENAME_TEMPLATE = os.environ.get('EXCEL_FILENAME_TEMPLATE', 'AWS_Health_Events_Analysis_{date}_{time}.xlsx')
customer_name = os.environ.get('CUSTOMER_NAME', 'Notification')
excluded_services_str = os.environ.get('EXCLUDED_SERVICES', '')
S3_BUCKET_NAME = os.environ.get('S3_BUCKET_NAME', '')
S3_KEY_PREFIX = os.environ.get('S3_KEY_PREFIX', '')
# Add these near the top with other environment variables
GROUP_ACCOUNT_IDS = os.environ.get('GROUP_ACCOUNT_IDS', '')
GROUP_EMAIL_IDS = os.environ.get('GROUP_EMAIL_IDS', '')
DYNAMODB_TABLE_NAME = os.environ.get('DYNAMODB_HEALTH_EVENTS_TABLE_NAME', '')


# NEW: Environment variable for specific accounts to create separate Excel files for
SPECIFIC_ACCOUNT_IDS = os.environ.get('SPECIFIC_ACCOUNT_IDS', '')
specific_account_ids = [s.strip() for s in SPECIFIC_ACCOUNT_IDS.split(',') if s.strip()]

excluded_services = [s.strip() for s in excluded_services_str.split(',') if s.strip()]

# Dictionary to store account ID to name mapping
account_id_to_name_map = {}

def get_account_name(account_id):
    """
    Get account name for a given account ID using AWS Organizations API
    
    Args:
        account_id (str): AWS account ID
        
    Returns:
        str: Account name or account ID if name can't be retrieved
    """
    # Check if we already have this account name in our cache
    if account_id in account_id_to_name_map:
        return account_id_to_name_map[account_id]
    
    try:
        # Try to get account name from Organizations API
        org_client = boto3.client('organizations')
        response = org_client.describe_account(AccountId=account_id)
        account_name = response.get('Account', {}).get('Name', account_id)
        
        # Cache the result
        account_id_to_name_map[account_id] = account_name
        return account_name
    except Exception as e:
        print(f"Error getting account name for {account_id}: {str(e)}")
        # If we can't get the name, just return the ID
        account_id_to_name_map[account_id] = account_id
        return account_id

def expand_events_by_account(events):
    """
    Expands events that affect multiple accounts into separate event records for each account.
    Fetches affected accounts if not already specified.
    
    Args:
        events (list): List of event dictionaries
        
    Returns:
        list: Expanded list of event dictionaries
    """
    expanded_events = []
    health_client = boto3.client('health', region_name='us-east-1')
    
    for event in events:
        # Get the account ID string which may contain multiple comma-separated IDs
        account_id_str = event.get('accountId', '')
        event_arn = event.get('arn', '')
        
        # If no account ID or it's N/A, try to fetch affected accounts
        if not account_id_str or account_id_str == 'N/A':
            try:
                print(f"Fetching affected accounts for event: {event.get('eventTypeCode', 'unknown')}")
                response = health_client.describe_affected_accounts_for_organization(
                    eventArn=event_arn
                )
                affected_accounts = response.get('affectedAccounts', [])
                
                if affected_accounts:
                    # If multiple accounts are affected, join them with commas
                    account_id_str = ', '.join(affected_accounts)
                    event['accountId'] = account_id_str # Update the event with the account IDs
                    print(f"Found affected accounts: {account_id_str}")
                else:
                    print("No affected accounts found")
                    expanded_events.append(event) # Keep the event as is
                    continue
            except Exception as e:
                print(f"Error fetching affected accounts: {str(e)}")
                expanded_events.append(event) # Keep the event as is
                continue
        
        # If no comma in the string, it's a single account or none
        if ',' not in account_id_str:
            expanded_events.append(event)
            continue
        
        # Split the account IDs and create a separate event for each
        account_ids = [aid.strip() for aid in account_id_str.split(',')]
        print(f"Expanding event {event_arn} for {len(account_ids)} accounts: {account_ids}")
        
        for account_id in account_ids:
            # Create a copy of the event for this specific account
            account_event = event.copy()
            account_event['accountId'] = account_id
            expanded_events.append(account_event)
    
    print(f"Expanded {len(events)} events to {len(expanded_events)} account-specific events")
    return expanded_events

def lambda_handler(event, context):
    print("Starting execution...")
    
    try:
        # Check if we're in single event processing mode
        single_event_arn = None
        if isinstance(event, dict) and 'event_arn' in event:
            single_event_arn = event['event_arn']
            print(f"Single event processing mode for ARN: {single_event_arn}")
        
        # Check if DynamoDB table name is provided
        DYNAMODB_TABLE_NAME = os.environ.get('DYNAMODB_HEALTH_EVENTS_TABLE_NAME', '')
        
        # If we have both single_event_arn and DYNAMODB_TABLE_NAME, process only that event
        if single_event_arn and DYNAMODB_TABLE_NAME:
            print(f"Processing single event {single_event_arn} for DynamoDB storage")
            
            # Initialize clients
            health_client = boto3.client('health', region_name='us-east-1')
            bedrock_client = get_bedrock_client()
            
            # Try to extract basic information from the ARN
            # ARN format: arn:aws:health:region::event/service/eventTypeCode/eventId
            arn_parts = single_event_arn.split('/')
            service = arn_parts[1] if len(arn_parts) > 1 else "UNKNOWN"
            event_type_code = arn_parts[2] if len(arn_parts) > 2 else "UNKNOWN"
            
            # Create a synthetic event with information from the ARN
            synthetic_event = {
                'arn': single_event_arn,
                'eventArn': single_event_arn,
                'eventTypeCode': event_type_code,
                'eventTypeCategory': 'issue', # Default category
                'service': service,
                'region': 'us-east-1', # Default region
                'startTime': datetime.utcnow().isoformat(),
                'lastUpdatedTime': datetime.utcnow().isoformat(),
                'accountId': 'N/A',
                'description': f"This is a synthetic event created for analysis of ARN: {single_event_arn}"
            }
            
            # Try to get more information from the events list
            try:
                use_org_view = is_org_view_enabled()
                print(f"Organization view enabled: {use_org_view}")
                
                # Create a filter based on the service from the ARN
                list_filter = {'services': [service]} if service != "UNKNOWN" else {}
                
                print(f"Attempting to list events with filter: {list_filter}")
                
                if use_org_view:
                    list_response = health_client.describe_events_for_organization(
                        filter=list_filter,
                        maxResults=100
                    )
                else:
                    list_response = health_client.describe_events(
                        filter=list_filter,
                        maxResults=100
                    )
                
                # Check if our event is in the list and update synthetic event
                if 'events' in list_response:
                    print(f"Found {len(list_response['events'])} events in list")
                    for evt in list_response['events']:
                        if evt.get('arn') == single_event_arn:
                            print(f"Found event in list, updating synthetic event with real data")
                            # Update our synthetic event with real data
                            synthetic_event.update({
                                'eventTypeCode': evt.get('eventTypeCode', event_type_code),
                                'eventTypeCategory': evt.get('eventTypeCategory', 'issue'),
                                'region': evt.get('region', 'us-east-1'),
                                'startTime': evt.get('startTime', synthetic_event['startTime']),
                                'lastUpdatedTime': evt.get('lastUpdatedTime', synthetic_event['lastUpdatedTime']),
                                'service': evt.get('service', service),
                                'statusCode': evt.get('statusCode', 'unknown')
                            })
                            break
                    else:
                        print(f"Event {single_event_arn} not found in the events list")
            except Exception as e:
                print(f"Error trying to get event from list: {str(e)}")
                traceback.print_exc()
            
            # Try to get affected accounts
            affected_accounts = []
            try:
                if use_org_view:
                    print("Attempting to get affected accounts")
                    accounts_response = health_client.describe_affected_accounts_for_organization(
                        eventArn=single_event_arn
                    )
                    affected_accounts = accounts_response.get('affectedAccounts', [])
                    print(f"Found affected accounts: {affected_accounts}")
            except Exception as e:
                print(f"Error getting affected accounts: {str(e)}")
            
            # If we found affected accounts, process for each account
            events_analysis = []
            
            if affected_accounts:
                for account_id in affected_accounts:
                    account_event = synthetic_event.copy()
                    account_event['accountId'] = account_id
                    account_event['accountName'] = get_account_name(account_id)
                    
                    # Try to get entity details for this account
                    try:
                        entities_response = health_client.describe_affected_entities_for_organization(
                            organizationEntityFilters=[
                                {
                                    'eventArn': single_event_arn,
                                    'awsAccountId': account_id
                                }
                            ]
                        )
                        
                        entities = entities_response.get('entities', [])
                        if entities:
                            affected_resources = ", ".join([e.get('entityValue', '') for e in entities if e.get('entityValue')])
                            account_event['affected_resources'] = affected_resources if affected_resources else "None specified"
                    except Exception as e:
                        print(f"Error getting affected entities for account {account_id}: {str(e)}")
                    
                    # Process this account's event
                    account_analysis = process_single_event(bedrock_client, account_event)
                    if account_analysis:
                        events_analysis.extend(account_analysis)
            else:
                # Process as a single event with no specific account
                print(f"Processing synthetic event with no specific account: {json.dumps(synthetic_event, default=str)}")
                single_analysis = process_single_event(bedrock_client, synthetic_event)
                if single_analysis:
                    events_analysis.extend(single_analysis)
            
            # Store the analyzed events in DynamoDB
            if events_analysis:
                storage_result = store_events_in_dynamodb(events_analysis)
                
                return {
                    'statusCode': 200,
                    'body': json.dumps({
                        'event_arn': single_event_arn,
                        'analyzed_events': len(events_analysis),
                        'affected_accounts': len(affected_accounts),
                        'stored_in_dynamodb': storage_result.get('stored', 0),
                        'updated_in_dynamodb': storage_result.get('updated', 0),
                        'failed_to_store': storage_result.get('failed', 0),
                        'synthetic': True
                    })
                }
            else:
                return {
                    'statusCode': 404,
                    'body': json.dumps({'error': f"Failed to analyze event: {single_event_arn}"})
                }
        
        # If we're here, we're in normal processing mode (either batch DynamoDB or standard email)
        # Get all configuration from environment variables
        analysis_window_days = int(os.environ['ANALYSIS_WINDOW_DAYS'])
        
        # Get event categories to process from environment variable
        event_categories_to_process = []
        if 'EVENT_CATEGORIES' in os.environ and os.environ['EVENT_CATEGORIES'].strip():
            event_categories_to_process = [cat.strip() for cat in os.environ['EVENT_CATEGORIES'].split(',')]
            print(f"Will only process these event categories: {event_categories_to_process}")
        else:
            print("No EVENT_CATEGORIES specified, will process all event categories")

        excluded_services_str = os.environ.get('EXCLUDED_SERVICES', '')
        excluded_services = [s.strip() for s in excluded_services_str.split(',') if s.strip()]
        
        if excluded_services:
            print(f"Excluding services from analysis: {excluded_services}")
        
        # Set up time range for filtering using environment variable
        bedrock_client = get_bedrock_client()
        end_time = datetime.utcnow()
        start_time = end_time - timedelta(days=analysis_window_days)
        
        print(f"Fetching events between {start_time} and {end_time}")
        
        # Format dates properly for the API
        formatted_start = start_time.strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + 'Z'
        formatted_end = end_time.strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + 'Z'
        
        # Initialize AWS Health client
        health_client = boto3.client('health', region_name='us-east-1')
        
        # Initialize variables for event collection
        all_events = []
        filtered_count = 0
        
        try:
            # Check if we should use organization view or account view
            use_org_view = is_org_view_enabled()
            
            if use_org_view:
                print("Using AWS Health Organization View")
                
                # CHANGE 1: Fetch closed events with both start and end date filters
                closed_filter = {
                    'startTime': {'from': formatted_start},
                    'endTime': {'to': formatted_end},
                    'eventStatusCodes': ['closed', 'upcoming']
                }
                
                # Add event type categories filter if specified
                if event_categories_to_process:
                    closed_filter['eventTypeCategories'] = event_categories_to_process
                
                print(f"Fetching CLOSED events with filter: {closed_filter}")
                closed_response = health_client.describe_events_for_organization(
                    filter=closed_filter,
                    maxResults=100
                )
                
                if 'events' in closed_response:
                    all_events.extend(closed_response['events'])
                    print(f"Retrieved {len(closed_response.get('events', []))} closed events")
                
                # Handle pagination for closed events
                while 'nextToken' in closed_response and closed_response['nextToken']:
                    print(f"Found nextToken for closed events, fetching more...")
                    if context.get_remaining_time_in_millis() < 15000: # 15 seconds buffer
                        print("Approaching Lambda timeout, stopping pagination")
                        break
                        
                    closed_response = health_client.describe_events_for_organization(
                        filter=closed_filter,
                        maxResults=100,
                        nextToken=closed_response['nextToken']
                    )
                    
                    if 'events' in closed_response:
                        all_events.extend(closed_response['events'])
                        print(f"Retrieved {len(closed_response.get('events', []))} additional closed events")
                
                # CHANGE 2: Fetch open events with only start date filter
                open_filter = {
                    'startTime': {'from': formatted_start}, # Started on or after start date
                    'eventStatusCodes': ['open'] # Only open events
                }
                
                # Add event type categories filter if specified
                if event_categories_to_process:
                    open_filter['eventTypeCategories'] = event_categories_to_process
                
                print(f"Fetching OPEN events with filter: {open_filter}")
                open_response = health_client.describe_events_for_organization(
                    filter=open_filter,
                    maxResults=100
                )
                
                if 'events' in open_response:
                    all_events.extend(open_response['events'])
                    print(f"Retrieved {len(open_response.get('events', []))} open events")
                
                # Handle pagination for open events
                while 'nextToken' in open_response and open_response['nextToken']:
                    print(f"Found nextToken for open events, fetching more...")
                    if context.get_remaining_time_in_millis() < 15000: # 15 seconds buffer
                        print("Approaching Lambda timeout, stopping pagination")
                        break
                        
                    open_response = health_client.describe_events_for_organization(
                        filter=open_filter,
                        maxResults=100,
                        nextToken=open_response['nextToken']
                    )
                    
                    if 'events' in open_response:
                        all_events.extend(open_response['events'])
                        print(f"Retrieved {len(open_response.get('events', []))} additional open events")
                
            else:
                print("Using AWS Health Account View")
                
                # CHANGE 3: Same approach for account view - fetch closed events
                closed_filter = {
                    'startTime': {'from': formatted_start},
                    'endTime': {'to': formatted_end},
                    'eventStatusCodes': ['closed', 'upcoming']
                }
                
                # Add event type categories filter if specified
                if event_categories_to_process:
                    closed_filter['eventTypeCategories'] = event_categories_to_process
                
                print(f"Fetching CLOSED events with filter: {closed_filter}")
                closed_response = health_client.describe_events(
                    filter=closed_filter,
                    maxResults=100
                )
                
                if 'events' in closed_response:
                    all_events.extend(closed_response['events'])
                
                # Handle pagination for closed events
                while 'nextToken' in closed_response and closed_response['nextToken']:
                    if context.get_remaining_time_in_millis() < 15000: # 15 seconds buffer
                        print("Approaching Lambda timeout, stopping pagination")
                        break
                        
                    closed_response = health_client.describe_events(
                        filter=closed_filter,
                        maxResults=100,
                        nextToken=closed_response['nextToken']
                    )
                    
                    if 'events' in closed_response:
                        all_events.extend(closed_response['events'])
                
                # CHANGE 4: Fetch open events with only start date filter
                open_filter = {
                    'startTime': {'from': formatted_start}, # Started on or after start date
                    'eventStatusCodes': ['open'] # Only open events
                }
                
                # Add event type categories filter if specified
                if event_categories_to_process:
                    open_filter['eventTypeCategories'] = event_categories_to_process
                
                print(f"Fetching OPEN events with filter: {open_filter}")
                open_response = health_client.describe_events(
                    filter=open_filter,
                    maxResults=100
                )
                
                if 'events' in open_response:
                    all_events.extend(open_response['events'])
                
                # Handle pagination for open events
                while 'nextToken' in open_response and open_response['nextToken']:
                    if context.get_remaining_time_in_millis() < 15000: # 15 seconds buffer
                        print("Approaching Lambda timeout, stopping pagination")
                        break
                        
                    open_response = health_client.describe_events(
                        filter=open_filter,
                        maxResults=100,
                        nextToken=open_response['nextToken']
                    )
                    
                    if 'events' in open_response:
                        all_events.extend(open_response['events'])
        
        except ClientError as e:
            if e.response['Error']['Code'] == 'SubscriptionRequiredException':
                print("Health Organization View is not enabled. Falling back to account-specific view.")
                
                # CHANGE 5: Same approach for fallback - fetch closed events
                closed_filter = {
                    'startTime': {'from': formatted_start},
                    'endTime': {'to': formatted_end},
                    'eventStatusCodes': ['closed', 'upcoming']
                }
                
                # Add event type categories filter if specified
                if event_categories_to_process:
                    closed_filter['eventTypeCategories'] = event_categories_to_process
                
                print(f"Fetching CLOSED events with filter: {closed_filter}")
                closed_response = health_client.describe_events(
                    filter=closed_filter,
                    maxResults=100
                )
                
                if 'events' in closed_response:
                    all_events.extend(closed_response['events'])
                
                # Handle pagination for closed events
                while 'nextToken' in closed_response and closed_response['nextToken']:
                    if context.get_remaining_time_in_millis() < 15000: # 15 seconds buffer
                        print("Approaching Lambda timeout, stopping pagination")
                        break
                        
                    closed_response = health_client.describe_events(
                        filter=closed_filter,
                        maxResults=100,
                        nextToken=closed_response['nextToken']
                    )
                    
                    if 'events' in closed_response:
                        all_events.extend(closed_response['events'])
                
                # CHANGE 6: Fetch open events with only start date filter
                open_filter = {
                    'startTime': {'from': formatted_start}, # Started on or after start date
                    'eventStatusCodes': ['open'] # Only open events
                }
                
                # Add event type categories filter if specified
                if event_categories_to_process:
                    open_filter['eventTypeCategories'] = event_categories_to_process
                
                print(f"Fetching OPEN events with filter: {open_filter}")
                open_response = health_client.describe_events(
                    filter=open_filter,
                    maxResults=100
                )
                
                if 'events' in open_response:
                    all_events.extend(open_response['events'])
                
                # Handle pagination for open events
                while 'nextToken' in open_response and open_response['nextToken']:
                    if context.get_remaining_time_in_millis() < 15000: # 15 seconds buffer
                        print("Approaching Lambda timeout, stopping pagination")
                        break
                        
                    open_response = health_client.describe_events(
                        filter=open_filter,
                        maxResults=100,
                        nextToken=open_response['nextToken']
                    )
                    
                    if 'events' in open_response:
                        all_events.extend(open_response['events'])
            else:
                raise
        
        # CHANGE 7: Remove duplicates by ARN
        unique_events = {}
        for item in all_events:
            arn = item.get('arn')
            if arn and arn not in unique_events:
                unique_events[arn] = item
        
        all_events = list(unique_events.values())

        # Filter out excluded services post-retrieval
        if excluded_services:
            filtered_events = [e for e in all_events if e.get('service') not in excluded_services]
            print(f"Filtered out {len(all_events) - len(filtered_events)} events from excluded services")
            all_events = filtered_events
        
        # NEW STEP: Expand events for multiple accounts
        all_events_original = all_events.copy()
        all_events_expanded = expand_events_by_account(all_events)
        print(f"Expanded {len(all_events_original)} events to {len(all_events_expanded)} account-specific events")
        
        items_count = len(all_events_original) # Keep the original count for reporting
        print(f"Fetched {items_count} unique events from AWS Health API (expanded to {len(all_events_expanded)} account-specific events)")
        
        if len(all_events_expanded) == 0:
            print("No events found with filter")
            return {
                'statusCode': 200,
                'body': json.dumps({
                    'message': 'No events matched the filter criteria',
                    'filters_used': {
                        'closed_filter': closed_filter if 'closed_filter' in locals() else {},
                        'open_filter': open_filter if 'open_filter' in locals() else {}
                    }
                })
            }
        
        # Process events directly from the API results
        events_analysis = []
        event_categories = defaultdict(int)
        raw_events = [] # Store raw event data for Excel
        
        # Process each event from the expanded API results
        for item in all_events_expanded:
            if context.get_remaining_time_in_millis() > 10000:
                # Check if we should process this event category
                event_type_category = item.get('eventTypeCategory', '')
                
                # Skip events that don't match our configured categories (this is redundant since we're filtering in the API call,
                # but keeping it for consistency with the original code)
                if event_categories_to_process and event_type_category not in event_categories_to_process:
                    print(f"Skipping event {item.get('eventTypeCode', 'unknown')} with category {event_type_category} (not in configured categories)")
                    filtered_count += 1
                    continue
                
                print(f"Processing event: {item.get('eventTypeCode', 'unknown')} with category {event_type_category}")
                
                try:
                    # Store raw event data for Excel
                    raw_events.append(item)
                    
                    # Ensure we have the event ARN and standardize field name
                    event_arn = item.get('arn', '')
                    if event_arn:
                        item['eventArn'] = event_arn # Standardize field name
                
                    # Extract account ID from ARN - this is already handled by the expansion function
                    account_id = item.get('accountId', 'N/A')
                    print(f"Processing with account ID: {account_id}")
                    
                    # NEW: Get account name
                    account_name = get_account_name(account_id)
                    
                    # Fetch additional details from Health API - now with a single account ID
                    health_data = fetch_health_event_details1(item.get('arn', ''), account_id)
                    
                    # Extract the actual description for analysis - IMPROVED EXTRACTION
                    actual_description = health_data['details'].get('eventDescription', {}).get('latestDescription', '')
                    
                    # If no description from Health API, try other possible fields
                    if not actual_description:
                        actual_description = (
                            item.get('eventDescription', '') or 
                            item.get('description', '') or 
                            item.get('message', '') or
                            'No description available'
                        )
                    
                    # Log the description we found
                    print(f"Using description (length: {len(actual_description)}): {actual_description[:100]}...")
                    
                    # Update the item with the actual description to improve analysis
                    item_with_description = item.copy()
                    item_with_description['description'] = actual_description
                    
                    analysis = analyze_event_with_bedrock(bedrock_client, item_with_description)
                    
                    categories = categorize_analysis(analysis)
                    if categories.get('critical', False):
                        event_categories['critical'] += 1
                    
                    risk_level = categories.get('risk_level', 'low')
                    event_categories[f"{risk_level}_risk"] += 1
                    
                    account_impact = categories.get('account_impact', 'low')
                    event_categories[f"{account_impact}_impact"] += 1
                    
                    # Create structured event data with both raw data and analysis
                    event_entry = {
                        "arn": item.get('arn', 'N/A'),
                        "eventArn": item.get('eventArn', item.get('arn', 'N/A')), # Ensure eventArn is included
                        "event_type": item.get('eventTypeCode', 'N/A'),
                        "description": actual_description,
                        "region": item.get('region', 'N/A'),
                        "start_time": format_time(item.get('startTime', 'N/A')),
                        "last_update_time": format_time(item.get('lastUpdatedTime', 'N/A')),
                        "event_type_category": item.get('eventTypeCategory', 'N/A'),
                        "analysis_text": analysis,
                        "critical": categories.get('critical', False),
                        "risk_level": categories.get('risk_level', 'low'),
                        "accountId": account_id, # Use the single account ID from the expanded event
                        "accountName": account_name, # NEW: Add account name
                        "impact_analysis": categories.get('impact_analysis', ''),
                        "required_actions": categories.get('required_actions', ''),
                        "time_sensitivity": categories.get('time_sensitivity', 'Routine'),
                        "risk_category": categories.get('risk_category', 'Unknown'),
                        "consequences_if_ignored": categories.get('consequences_if_ignored', ''),
                        "affected_resources": extract_affected_resources(health_data['entities']),
                        # NEW: Add event impact type
                        "event_impact_type": categories.get('event_impact_type', 'Unknown')
                    }
                    
                    events_analysis.append(event_entry)
                    print(f"Successfully analyzed event {len(events_analysis)}")
                except Exception as e:
                    print(f"Error analyzing event: {str(e)}")
                    traceback.print_exc()
            else:
                print("Approaching Lambda timeout, stopping event processing")
                break
        
        if events_analysis:
            print(f"Successfully analyzed {len(events_analysis)} events (filtered out {filtered_count} events)")
            
            # Check if we should store in DynamoDB instead of sending emails
            if DYNAMODB_TABLE_NAME:
                print(f"DynamoDB table name provided: {DYNAMODB_TABLE_NAME}")
                print("Storing events in DynamoDB instead of sending emails")
                
                # Store events in DynamoDB
                storage_result = store_events_in_dynamodb(events_analysis)
                
                try:
                    # Add CloudWatch metrics with error handling
                    add_cloudwatch_metrics(event_categories, len(events_analysis), items_count, filtered_count)
                except Exception as e:
                    print(f"Error publishing CloudWatch metrics: {str(e)}")
                    print("Continuing execution despite metrics error")
                
                return {
                    'statusCode': 200,
                    'body': json.dumps({
                        'total_events': items_count,
                        'total_expanded_events': len(all_events_expanded),
                        'analyzed_events': len(events_analysis),
                        'filtered_events': filtered_count,
                        'stored_in_dynamodb': storage_result['stored'],
                        'updated_in_dynamodb': storage_result.get('updated', 0),
                        'failed_to_store': storage_result['failed'],
                        'categories': dict(event_categories)
                    })
                }
            else:
                                # Create a complete Excel report with all events
                complete_excel_buffer = create_excel_report_improved_single_sheet(events_analysis)
                
                # Generate filename for complete report
                current_date = datetime.now().strftime('%Y-%m-%d')
                current_time = datetime.now().strftime('%H-%M-%S')
                complete_excel_filename = EXCEL_FILENAME_TEMPLATE.format(date=current_date, time=current_time)
                
                # Generate summary HTML with filtering info for complete report
                complete_summary_html = generate_summary_html(
                    items_count,  # Use original count before expansion
                    event_categories, 
                    filtered_count, 
                    event_categories_to_process if event_categories_to_process else None,
                    events_analysis 
                )
                
                # Always send the complete report to default recipients and upload to S3
                send_ses_email_with_attachment(
                    complete_summary_html, 
                    complete_excel_buffer, 
                    items_count, 
                    event_categories, 
                    events_analysis,
                    complete_excel_filename
                )
                
                # Parse group configurations
                account_groups = parse_group_config(GROUP_ACCOUNT_IDS)
                email_groups = parse_group_config(GROUP_EMAIL_IDS)
                
                # Check if we should use group-based processing
                use_group_processing = validate_group_configs(account_groups, email_groups)
                
                if use_group_processing:
                    print("Using group-based processing")
                    
                    # Group events by account groups
                    group_events = group_events_by_account_groups(events_analysis, account_groups)
                    
                    # Create Excel reports for each group
                    for group_name, events in group_events.items():
                        if events and group_name != 'other':  # Skip empty groups and "other" group
                            # Create safe group name for filename (replace spaces with underscores)
                            safe_group_name = group_name.replace(' ', '_')
                            
                            # Create Excel report for this group
                            excel_buffer = create_excel_report_improved_single_sheet(events)
                            filename = f"AWS_Health_Events_Analysis_{safe_group_name}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
                            
                            # Get email recipients for this group
                            group_recipients = email_groups.get(group_name, [])
                            
                            if group_recipients:
                                # Send email with this group's Excel file
                                send_ses_email_with_group_attachment(
                                    complete_summary_html,
                                    excel_buffer,
                                    filename,
                                    items_count,
                                    filtered_count,
                                    event_categories_to_process if event_categories_to_process else None,
                                    events,
                                    group_recipients
                                )
                
                elif specific_account_ids:
                    # Fall back to SPECIFIC_ACCOUNT_IDS logic
                    print("Group configuration invalid or not present. Falling back to SPECIFIC_ACCOUNT_IDS logic.")
                    
                    # Create a dictionary to hold events for each specific account
                    account_events = {account_id: [] for account_id in specific_account_ids}
                    # Add a key for "other" accounts
                    account_events['other'] = []
                    
                    # Distribute events to their respective accounts
                    for event in events_analysis:
                        account_id = event.get('accountId', 'N/A')
                        if account_id in specific_account_ids:
                            account_events[account_id].append(event)
                        else:
                            account_events['other'].append(event)
                    
                    # Create Excel reports for each account
                    excel_buffers = {}
                    excel_filenames = {}
                    
                    for account_id, events in account_events.items():
                        if events:  # Only create reports for accounts with events
                            if account_id == 'other':
                                account_name = 'Other_Accounts'
                            else:
                                account_name = get_account_name(account_id).replace(' ', '_')
                            
                            excel_buffer = create_excel_report_improved_single_sheet(events)
                            filename = f"AWS_Health_Events_Analysis_{account_name}_{account_id}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
                            excel_buffers[account_id] = excel_buffer
                            excel_filenames[account_id] = filename
                    
                    # Send a single email with all account-specific attachments
                    send_ses_email_with_multiple_attachments(
                        complete_summary_html, 
                        excel_buffers,
                        excel_filenames,
                        items_count, 
                        event_categories, 
                        events_analysis
                    )
                
                try:
                    # Add CloudWatch metrics with error handling
                    add_cloudwatch_metrics(event_categories, len(events_analysis), items_count, filtered_count)
                except Exception as e:
                    print(f"Error publishing CloudWatch metrics: {str(e)}")
                    print("Continuing execution despite metrics error")

            
            return {
                'statusCode': 200,
                'body': json.dumps({
                    'total_events': items_count, # Original event count
                    'total_expanded_events': len(all_events_expanded), # Expanded event count
                    'analyzed_events': len(events_analysis),
                    'filtered_events': filtered_count,
                    'categories': dict(event_categories),
                    'category_filter_applied': bool(event_categories_to_process),
                    'categories_processed': event_categories_to_process,
                    'filters_used': {
                        'closed_filter': closed_filter if 'closed_filter' in locals() else {},
                        'open_filter': open_filter if 'open_filter' in locals() else {}
                    }
                })
            }
        else:
            print(f"No events were successfully analyzed (filtered out {filtered_count} events)")
            return {
                'statusCode': 200,
                'body': json.dumps({
                    'message': 'Found events but none were analyzed',
                    'events_found': items_count,
                    'expanded_events_found': len(all_events_expanded),
                    'filtered_events': filtered_count,
                    'category_filter_applied': bool(event_categories_to_process),
                    'categories_processed': event_categories_to_process,
                    'filters_used': {
                        'closed_filter': closed_filter if 'closed_filter' in locals() else {},
                        'open_filter': open_filter if 'open_filter' in locals() else {}
                    }
                })
            }
            
    except Exception as e:
        print(f"Error: {str(e)}")
        traceback.print_exc()
        return {
            'statusCode': 500,
            'body': json.dumps({'error': str(e)})
        }


def is_org_view_enabled():
    """
    Check if AWS Health Organization View is enabled
    
    Returns:
        bool: True if organization view is enabled, False otherwise
    """
    try:
        # Try to call an organization-specific API to check if it's enabled
        health_client = boto3.client('health', region_name='us-east-1')
        # This will throw an exception if org view is not enabled
        health_client.describe_events_for_organization(
            filter={},
            maxResults=1
        )
        return True
    except Exception as e:
        error_code = getattr(e, 'response', {}).get('Error', {}).get('Code', '')
        if error_code == 'SubscriptionRequiredException':
            return False
        # For any other error, assume we don't have org view permissions
        return False

def get_bedrock_client():
    """
    Get Amazon Bedrock client
    
    Returns:
        boto3.client: Bedrock runtime client
    """
    return boto3.client(service_name='bedrock-runtime',region_name='us-east-1')

def format_time(time_str):
    """
    Format time string to be consistent
    
    Args:
        time_str (str): ISO format time string
        
    Returns:
        str: Formatted time string (YYYY-MM-DD)
    """
    if not time_str or time_str == 'N/A':
        return 'N/A'
    
    try:
        # If it's already a datetime object
        if isinstance(time_str, datetime):
            return time_str.strftime('%Y-%m-%d')
        
        # Parse ISO format
        dt = datetime.fromisoformat(time_str.replace('Z', '+00:00'))
        return dt.strftime('%Y-%m-%d')
    except Exception:
        # If we can't parse it, return as is
        return time_str

def fetch_health_event_details(event_arn):
    """
    Fetch detailed event information from AWS Health API
    
    Args:
        event_arn (str): ARN of the health event
        
    Returns:
        dict: Event details including affected resources
    """
    try:
        health_client = boto3.client('health', region_name='us-east-1')
        
        # Get event details
        event_details = health_client.describe_event_details(
            eventArns=[event_arn]
        )
        
        # Get affected entities
        affected_entities = health_client.describe_affected_entities(
            filter={
                'eventArns': [event_arn]
            }
        )
        
        return {
            'details': event_details.get('successfulSet', [{}])[0] if event_details.get('successfulSet') else {},
            'entities': affected_entities.get('entities', [])
        }
    except Exception as e:
        print(f"Error fetching Health API data: {str(e)}")
        return {'details': {}, 'entities': []}
         
def extract_affected_resources(entities):
    """
    Extract affected resources from Health API entities
    
    Args:
        entities (list): List of entity objects from Health API
        
    Returns:
        str: Comma-separated list of affected resources
    """
    if not entities:
        return "None specified"
    
    resources = []
    for entity in entities:
        entity_value = entity.get('entityValue', '')
        if entity_value:
            resources.append(entity_value)
    
    if resources:
        return ", ".join(resources)
    else:
        return "None specified"

def analyze_event_with_bedrock(bedrock_client, event_data):
    """
    Analyze an AWS Health event using Amazon Bedrock with focus on outage impact
    
    Args:
        bedrock_client: Amazon Bedrock client
        event_data (dict): Event data to analyze
        
    Returns:
        dict: Analyzed event data
    """
    try:
        # Get event details
        event_type = event_data.get('eventTypeCode', event_data.get('event_type', 'Unknown'))
        event_category = event_data.get('eventTypeCategory', event_data.get('event_type_category', 'Unknown'))
        region = event_data.get('region', 'Unknown')
        
        # Format start time if it's a datetime object
        start_time = event_data.get('startTime', event_data.get('start_time', 'Unknown'))
        if hasattr(start_time, 'isoformat'):
            start_time = start_time.isoformat()
        
        # Use description for analysis
        description = event_data.get('description', 'No description available')

        
        # Prepare prompt for Bedrock - ENHANCED FOR OUTAGE ANALYSIS AND EVENT IMPACT TYPE
        print(f"Processing event: {event_type} with category {event_category}")
        print(f"Using description (length: {len(description)}): {description[:100]}...")
        
        prompt = f"""
        You are an AWS expert specializing in outage analysis and business continuity. Your task is to analyze this AWS Health event and determine its potential impact on workload availability, system connectivity, and service outages.
        
        AWS Health Event:
        - Type: {event_type}
        - Category: {event_category}
        - Region: {region}
        - Start Time: {start_time}
        
        Event Description:
        {description}
        
        IMPORTANT ANALYSIS FOCUS:
        1. Will this event cause workload downtime if required actions are not taken?
        2. Will there be any service outages associated with this event?
        3. Will the application/workload experience network integration issues between connecting systems?
        4. What specific AWS services or resources could be impacted?
         
        
        CRITICAL EVENT CRITERIA:
        - Any event that will cause service downtime should be marked as CRITICAL
        - Any event that will cause network integration or SSL issues between systems should be marked as CRITICAL
        - Any event that requires immediate action to prevent outage should be marked as URGENT time sensitivity
        - Events with high impact but no immediate downtime should be marked as HIGH risk level
    
        Please analyze this event and provide the following information in JSON format:
        {{
          "critical": boolean,
          "risk_level": "critical|high|medium|low",
          "account_impact": "critical|high|medium|low",
          "time_sensitivity": "Routine|Urgent|Critical",
          "risk_category": "Availability|Security|Performance|Cost|Compliance",
          "required_actions": "string",
          "impact_analysis": "string",
          "consequences_if_ignored": "string",
          "affected_resources": "string",
          "event_impact_type": "Service Outage|Billing Impact|Security Issue|Performance Degradation|Maintenance|Informational"
        }}
        
        IMPORTANT: In your impact_analysis field, be very specific about:
        1. Potential outages and their estimated duration
        2. Connectivity issues between systems
        3. Whether this will cause downtime if actions are not taken
        
        In your consequences_if_ignored field, clearly state what outages or disruptions will occur if the event is not addressed.

        RISK LEVEL GUIDELINES:
        - CRITICAL: Will cause service outage or severe disruption if not addressed
        - HIGH: Significant impact but not an immediate outage
        - MEDIUM: Moderate impact requiring attention
        - LOW: Minimal impact, routine maintenance
        
        EVENT IMPACT TYPE GUIDELINES:
        - Service Outage: Event will cause or is causing service unavailability
        - Billing Impact: Event affects billing or costs
        - Security Issue: Event relates to security vulnerabilities or threats
        - Performance Degradation: Event causes reduced performance but not complete outage
        - Maintenance: Planned maintenance with minimal impact
        - Informational: General information with no direct impact
        """
        
        # Determine which model we're using and format accordingly
        model_id = os.environ.get('BEDROCK_MODEL_ID', 'anthropic.claude-v2')
        max_tokens = int(os.environ.get('BEDROCK_MAX_TOKENS', '4000'))
        temperature = float(os.environ.get('BEDROCK_TEMPERATURE', '0.2'))
        top_p = float(os.environ.get('BEDROCK_TOP_P', '0.9'))
        
        print(f"Sending request to Bedrock model: '{model_id}'")
        
        if "claude-3" in model_id.lower():
            # Claude 3 models use the messages format
            payload = {
                "modelId": model_id,
                "contentType": "application/json",
                "accept": "application/json",
                "body": json.dumps({
                    "anthropic_version": "bedrock-2023-05-31",
                    "max_tokens": max_tokens,
                    "temperature": temperature,
                    "top_p": top_p,
                    "messages": [
                        {
                            "role": "user",
                            "content": prompt
                        }
                    ]
                })
            }
        else:
            # Claude 2 and other models use the older prompt format
            payload = {
                "modelId": model_id,
                "contentType": "application/json",
                "accept": "application/json",
                "body": json.dumps({
                    "prompt": f"\n\nHuman: {prompt}\n\nAssistant:",
                    "max_tokens_to_sample": max_tokens,
                    "temperature": temperature,
                    "top_p": top_p
                })
            }
        
        # Call Bedrock
        try:
            response = bedrock_client.invoke_model(**payload)
            response_body = json.loads(response.get('body').read())
            
            # Extract response based on model
            if "claude-3" in model_id.lower():
                response_text = response_body.get('content', [{}])[0].get('text', '')
            else:
                response_text = response_body.get('completion', '')
            
            # Store the full analysis text as a string
            event_data['analysis_text'] = response_text
            
            # Try to extract JSON from the response
            json_match = re.search(r'```json\s*(.*?)\s*```', response_text, re.DOTALL)
            if json_match:
                json_str = json_match.group(1)
            else:
                json_match = re.search(r'({.*})', response_text, re.DOTALL)
                if json_match:
                    json_str = json_match.group(1)
                else:
                    json_str = response_text
            
            # Parse the JSON
            try:
                analysis = json.loads(json_str)
                # Normalize risk level to ensure consistency
                if 'risk_level' in analysis:
                    risk_level = analysis['risk_level'].strip().upper()
                    
                    # Ensure "critical" is properly recognized and distinguished from "high"
                    if risk_level in ['CRITICAL', 'SEVERE']:
                        analysis['risk_level'] = 'CRITICAL'
                        # Make sure critical boolean flag is consistent
                        analysis['critical'] = True
                    elif risk_level == 'HIGH':
                        analysis['risk_level'] = 'HIGH'
                    elif risk_level in ['MEDIUM', 'MODERATE']:
                        analysis['risk_level'] = 'MEDIUM'
                    elif risk_level == 'LOW':
                        analysis['risk_level'] = 'LOW'
                    
                    # If critical flag is True but risk_level isn't CRITICAL, fix it
                    if analysis.get('critical', False) and analysis['risk_level'] != 'CRITICAL':
                        analysis['risk_level'] = 'CRITICAL'
                
                # Update event data with analysis
                event_data.update(analysis)
                
                return event_data
            except json.JSONDecodeError:
                print(f"Failed to parse JSON from response: {response_text[:200]}...")
                # Provide default values if parsing fails
                event_data.update({
                    'critical': False,
                    'risk_level': 'low',
                    'account_impact': 'low',
                    'time_sensitivity': 'Routine',
                    'risk_category': 'Unknown',
                    'required_actions': 'Review event details manually',
                    'impact_analysis': 'Unable to automatically analyze this event',
                    'consequences_if_ignored': 'Unknown',
                    'affected_resources': 'Unknown',
                    'event_impact_type': 'Informational'
                })
                return event_data
                
        except Exception as e:
            print(f"Error in Bedrock analysis: {str(e)}")
            traceback.print_exc()
            
            # Provide default values if Bedrock analysis fails
            event_data.update({
                'critical': False,
                'risk_level': 'low',
                'account_impact': 'low',
                'time_sensitivity': 'Routine',
                'risk_category': 'Unknown',
                'required_actions': 'Review event details manually',
                'impact_analysis': 'Unable to automatically analyze this event',
                'consequences_if_ignored': 'Unknown',
                'affected_resources': 'Unknown',
                'analysis_text': f"Error during analysis: {str(e)}",
                'event_impact_type': 'Informational'
            })
            return event_data
    
    except Exception as e:
        print(f"Unexpected error in analyze_event_with_bedrock: {str(e)}")
        traceback.print_exc()
        
        # Provide default values if function fails
        event_data.update({
            'critical': False,
            'risk_level': 'low',
            'account_impact': 'low',
            'time_sensitivity': 'Routine',
            'risk_category': 'Unknown',
            'required_actions': 'Review event details manually',
            'impact_analysis': 'Unable to automatically analyze this event',
            'consequences_if_ignored': 'Unknown',
            'affected_resources': 'Unknown',
            'analysis_text': f"Error during analysis: {str(e)}",
            'event_impact_type': 'Informational'
        })
        return event_data

def categorize_analysis(analysis_text):
    """
    Extract structured data from Bedrock analysis text
    
    Args:
        analysis_text: Analysis text from Bedrock (string or dict)
        
    Returns:
        dict: Structured data extracted from analysis
    """
    categories = {
        'critical': False,
        'risk_level': 'low',
        'impact_analysis': '',
        'required_actions': '',
        'time_sensitivity': 'Routine',
        'risk_category': 'Unknown',
        'consequences_if_ignored': '',
        'event_category': 'Low',
        'event_impact_type': 'Informational' # Added default value for new field
    }
    
    try:
        # If analysis_text is already a dictionary, use it directly
        if isinstance(analysis_text, dict):
            # Update our categories with values from the dictionary
            for key in categories.keys():
                if key in analysis_text:
                    categories[key] = analysis_text[key]
            
            # Also check for affected_resources
            if 'affected_resources' in analysis_text:
                categories['affected_resources'] = analysis_text['affected_resources']
                
            return categories
            
        # If analysis_text is not a string, convert it to string
        if not isinstance(analysis_text, str):
            analysis_text = str(analysis_text)
            
        # Try to parse as JSON first
        try:
            json_data = json.loads(analysis_text)
            # If successful, update our categories with values from the JSON
            for key in categories.keys():
                if key in json_data:
                    categories[key] = json_data[key]
            
            # Also check for affected_resources
            if 'affected_resources' in json_data:
                categories['affected_resources'] = json_data['affected_resources']
                
            return categories
        except json.JSONDecodeError:
            # Not valid JSON, continue with regex parsing
            pass
            
        # Extract critical status
        critical_match = re.search(r'CRITICAL:\s*(?:\[)?([Yy]es|[Nn]o)(?:\])?', analysis_text)
        if critical_match:
            categories['critical'] = critical_match.group(1).lower() == 'yes'
        
        # Extract risk level
        risk_match = re.search(r'RISK LEVEL:\s*(?:\[)?([Hh]igh|[Mm]edium|[Ll]ow)(?:\])?', analysis_text)
        if risk_match:
            categories['risk_level'] = risk_match.group(1).lower()
        
        # Extract account impact
        impact_match = re.search(r'ACCOUNT IMPACT:\s*(?:\[)?([Hh]igh|[Mm]edium|[Ll]ow)(?:\])?', analysis_text)
        if impact_match:
            categories['account_impact'] = impact_match.group(1).lower()
        
        # Extract impact analysis
        impact_analysis_match = re.search(r'IMPACT ANALYSIS:(.*?)(?:REQUIRED ACTIONS:|$)', analysis_text, re.DOTALL)
        if impact_analysis_match:
            categories['impact_analysis'] = impact_analysis_match.group(1).strip()
        
        # Extract required actions
        required_actions_match = re.search(r'REQUIRED ACTIONS:(.*?)(?:TIME SENSITIVITY:|$)', analysis_text, re.DOTALL)
        if required_actions_match:
            categories['required_actions'] = required_actions_match.group(1).strip()
        
        # Extract time sensitivity
        time_sensitivity_match = re.search(r'TIME SENSITIVITY:\s*([Ii]mmediate|[Uu]rgent|[Ss]oon|[Rr]outine)', analysis_text)
        if time_sensitivity_match:
            categories['time_sensitivity'] = time_sensitivity_match.group(1).capitalize()
        
        # Extract risk category
        risk_category_match = re.search(r'RISK CATEGORY:\s*([Tt]echnical|[Oo]perational|[Ss]ecurity|[Cc]ompliance|[Cc]ost|[Aa]vailability)', analysis_text)
        if risk_category_match:
            categories['risk_category'] = risk_category_match.group(1).capitalize()
        
        # Extract consequences if ignored
        consequences_match = re.search(r'CONSEQUENCES IF IGNORED:(.*?)(?:$)', analysis_text, re.DOTALL)
        if consequences_match:
            categories['consequences_if_ignored'] = consequences_match.group(1).strip()
        
        # Extract affected resources
        affected_match = re.search(r'AFFECTED RESOURCES:(.*?)(?:$)', analysis_text, re.DOTALL)
        if affected_match:
            categories['affected_resources'] = affected_match.group(1).strip()
            
        # Extract event impact type (new)
        event_impact_match = re.search(r'EVENT IMPACT TYPE:\s*(Service Outage|Billing Impact|Security Issue|Performance Degradation|Maintenance|Informational)', analysis_text)
        if event_impact_match:
            categories['event_impact_type'] = event_impact_match.group(1)
            
    except Exception as e:
        print(f"Error categorizing analysis: {str(e)}")
    
    return categories

def create_excel_report_improved_single_sheet(events_analysis):
    """
    Create an improved Excel report with only one sheet (All Events)
    
    Args:
        events_analysis (list): List of analyzed events data
        
    Returns:
        BytesIO: Excel file as bytes
    """
    # Create workbook and sheet
    wb = Workbook()
    events_sheet = wb.active
    events_sheet.title = "All Events"
    
    # Add headers to events sheet
    headers = [
        "Event ARN",
        "Event Type", 
        "Region", 
        "Start Time", 
        "Last Update", 
        "Category",
        "Description",
        "Critical", 
        "Risk Level", 
        "Account ID",
        "Account Name", # NEW: Added Account Name column
        "Time Sensitivity", 
        "Risk Category",
        "Event Impact Type", # NEW: Added Event Impact Type column
        "Required Actions", 
        "Impact Analysis", 
        "Consequences If Ignored", 
        "Affected Resources"
    ]
    
    for col_num, header in enumerate(headers, 1):
        cell = events_sheet.cell(row=1, column=col_num)
        cell.value = header
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    
    # Add event data
    for row_num, event in enumerate(events_analysis, 2):
        event_arn = event.get('eventArn', event.get('arn', 'N/A'))
        events_sheet.cell(row=row_num, column=1).value = event_arn
        events_sheet.cell(row=row_num, column=2).value = event.get('event_type', 'N/A')
        events_sheet.cell(row=row_num, column=3).value = event.get('region', 'N/A')
        events_sheet.cell(row=row_num, column=4).value = event.get('start_time', 'N/A')
        events_sheet.cell(row=row_num, column=5).value = event.get('last_update_time', 'N/A')
        events_sheet.cell(row=row_num, column=6).value = event.get('event_type_category', 'N/A')
        
        # Add description with text wrapping
        description_cell = events_sheet.cell(row=row_num, column=7)
        description_cell.value = event.get('description', 'N/A')
        description_cell.alignment = Alignment(wrap_text=True, vertical='top')
        
        events_sheet.cell(row=row_num, column=8).value = "Yes" if event.get('critical', False) else "No"
        events_sheet.cell(row=row_num, column=9).value = event.get('risk_level', 'low').upper()
        events_sheet.cell(row=row_num, column=10).value = event.get('accountId', 'N/A')
        events_sheet.cell(row=row_num, column=11).value = event.get('accountName', 'N/A') # NEW: Account Name
        events_sheet.cell(row=row_num, column=12).value = event.get('time_sensitivity', 'Routine')
        events_sheet.cell(row=row_num, column=13).value = event.get('risk_category', 'Unknown')
        events_sheet.cell(row=row_num, column=14).value = event.get('event_impact_type', 'Informational') # NEW: Event Impact Type
        events_sheet.cell(row=row_num, column=15).value = event.get('required_actions', '')
        events_sheet.cell(row=row_num, column=16).value = event.get('impact_analysis', '')
        events_sheet.cell(row=row_num, column=17).value = event.get('consequences_if_ignored', '')
        events_sheet.cell(row=row_num, column=18).value = event.get('affected_resources', 'None')
        
        # Color coding based on risk level
        risk_level = event.get('risk_level', 'low').lower()
        is_critical = event.get('critical', False)
        
        if is_critical:
            for col_num in range(1, 19): # Updated column range
                events_sheet.cell(row=row_num, column=col_num).fill = PatternFill(
                    start_color="FFCCCC", end_color="FFCCCC", fill_type="solid"
                )
        elif risk_level == 'high':
            for col_num in range(1, 19): # Updated column range
                events_sheet.cell(row=row_num, column=col_num).fill = PatternFill(
                    start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
                )
        elif risk_level == 'medium':
            for col_num in range(1, 19): # Updated column range
                events_sheet.cell(row=row_num, column=col_num).fill = PatternFill(
                    start_color="E6F2FF", end_color="E6F2FF", fill_type="solid"
                )
    
    # Auto-adjust column widths
    for col in events_sheet.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            if cell.value:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = min(len(str(cell.value)), 50) # Cap at 50 characters
                except:
                    pass
        adjusted_width = max_length + 2
        events_sheet.column_dimensions[column].width = adjusted_width
    
    # Set specific width for description column
    events_sheet.column_dimensions['G'].width = 60 # Description column
    events_sheet.column_dimensions['P'].width = 60 # Impact Analysis column
    events_sheet.column_dimensions['Q'].width = 60 # Consequences column
    
    # Save to BytesIO
    excel_buffer = BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    
    return excel_buffer

def generate_summary_html(total_events, event_categories, filtered_events, category_filter, events_analysis):
    """
    Generate HTML summary for email
    
    Args:
        total_events (int): Total number of events
        event_categories (dict): Event categories count
        filtered_events (int): Number of filtered events
        category_filter (list): Categories used for filtering
        events_analysis (list): Analyzed events data
        
    Returns:
        str: HTML content for email
    """
    # Current date for the report
    current_date = datetime.now().strftime('%Y-%m-%d')
    
    # Calculate accurate event counts directly from events_analysis
    critical_count = sum(1 for event in events_analysis if event.get('critical', False))
    high_risk_count = sum(1 for event in events_analysis if event.get('risk_level', '').lower() == 'high')
    medium_risk_count = sum(1 for event in events_analysis if event.get('risk_level', '').lower() == 'medium')
    low_risk_count = sum(1 for event in events_analysis if event.get('risk_level', '').lower() == 'low')

    # Define CSS for better table formatting
    table_css = """
    <style>
        .health-events-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 5px;
        }
        .health-events-table th, .health-events-table td {
            border: 1px solid #ddd;
            padding: 6px;
            text-align: left;
            vertical-align: top;
            word-wrap: break-word;
            font-size: 5px;
        }
        .health-events-table th {
            background-color: #f2f2f2;
        }
        /* Column width constraints */
        .col-arn {
            width: 15%;
        }
        .col-region {
            width: 8%;
        }
        .col-start-time {
            width: 12%;
        }
        .col-risk {
            width: 8%;
        }
        .col-accountid {
            width: 12%;
        }
    </style>
    """
    
    # Start building HTML content
    html_content = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; }}
            .header {{ background-color: #232F3E; color: white; padding: 20px; }}
            .content {{ padding: 20px; }}
            table {{ border-collapse: collapse; width: 100%; margin-top: 20px; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left;"font-size:5px;" }}
            th {{ background-color: #f2f2f2; }}
            .critical {{ background-color: #ffcccc; }}
            .high {{ background-color: #fff2cc; }}
            .medium {{ background-color: #e6f2ff; }}
            .summary {{ margin-top: 20px; margin-bottom: 20px; }}
        </style>
    </head>
    <body>
        <div class="header">
            <h1>AWS Health Events Analysis Report</h1>
            <p>Date: {current_date}</p>
        </div>
        <div class="content">
            <div class="summary">
                <h2>Summary</h2>
                <p>Total AWS Health events analyzed: {len(events_analysis)} of {total_events} events found</p>
    """
    
    # Add analysis window information
    end_time = datetime.utcnow()
    start_time = end_time - timedelta(days=int(os.environ['ANALYSIS_WINDOW_DAYS']))
    html_content += f"""
                <p>Analysis Window: {start_time.strftime('%Y-%m-%d %H:%M:%S')} UTC to {end_time.strftime('%Y-%m-%d %H:%M:%S')} UTC</p>
    """
    
    # Add filter information if applicable
    if category_filter:
        html_content += f"""
                <p>Events filtered by categories: {', '.join(category_filter)}</p>
                <p>Events excluded by filter: {filtered_events}</p>
        """
    
    # Add event category counts - USING ACCURATE COUNTS
    html_content += """
                <h3>Event Categories</h3>
                <ul>
    """
    
    # Add critical events count if any
    if critical_count > 0:
        html_content += f"""
                <li><strong>Critical Events:</strong> {critical_count}</li>
        """
    
    # Add risk level counts - USING ACCURATE COUNTS
    html_content += f"""
                <li><strong>High Risk Events:</strong> {high_risk_count}</li>
                <li><strong>Medium Risk Events:</strong> {medium_risk_count}</li>
                <li><strong>Low Risk Events:</strong> {low_risk_count}</li>
            </ul>
            </div>
    """
    
    # Add critical events table if any exist
    critical_events = [event for event in events_analysis if event.get('critical', False)]
    if critical_events:
        html_content += """
            <h2>Critical Events</h2>
            <table class="health-events-table">
                <tr>
                    <th style="font-size:12px;">Event ARN</th>
                    <th style="font-size:12px;">Region</th>
                    <th style="font-size:12px;">Start Time</th>
                    <th style="font-size:12px;">Risk Level</th>
                    <th style="font-size:12px;">Account ID</th>
                    <th style="font-size:12px;">Account Name</th>
                    <th style="font-size:12px;">Event Impact Type</th>
                </tr>
        """
        
        for event in critical_events:
            # Get event ARN, preferring eventArn if available, falling back to arn
            event_arn = event.get('eventArn', event.get('arn', 'N/A'))
            
            html_content += f"""
                <tr class="critical">
                    <td style="font-size:12px;">{event_arn}</td>
                    <td style="font-size:12px;">{event.get('region', 'N/A')}</td>
                    <td style="font-size:12px;">{event.get('start_time', 'N/A')}</td>
                    <td style="font-size:12px;">{event.get('risk_level', 'N/A').upper()}</td>
                    <td style="font-size:12px;">{event.get('accountId', 'N/A')}</td>
                    <td style="font-size:12px;">{event.get('accountName', 'N/A')}</td>
                    <td style="font-size:12px;">{event.get('event_impact_type', 'Informational')}</td>
                </tr>
            """
        
        html_content += """
            </table>
        """
    
    # Add high risk events table
    high_risk_events = [event for event in events_analysis if event.get('risk_level', '').lower() == 'high']
    if high_risk_events:
        html_content += """
            <h2>High Risk Events</h2>
            <table class="health-events-table">
                <tr>
                    <th style="font-size:12px;">Event ARN</th>
                    <th style="font-size:12px;">Region</th>
                    <th style="font-size:12px;">Start Time</th>
                    <th style="font-size:12px;">Risk Level</th>
                    <th style="font-size:12px;">Account ID</th>
                    <th style="font-size:12px;">Account Name</th>
                    <th style="font-size:12px;">Event Impact Type</th>
                </tr>
        """
        
        for event in high_risk_events:
            # Get event ARN, preferring eventArn if available, falling back to arn
            event_arn = event.get('eventArn', event.get('arn', 'N/A'))
            
            html_content += f"""
                <tr class="high">
                    <td style="font-size:12px;">{event_arn}</td>
                    <td style="font-size:12px;">{event.get('region', 'N/A')}</td>
                    <td style="font-size:12px;">{event.get('start_time', 'N/A')}</td>
                    <td style="font-size:12px;">{event.get('risk_level', 'N/A').upper()}</td>
                    <td style="font-size:12px;">{event.get('accountId', 'N/A')}</td>
                    <td style="font-size:12px;">{event.get('accountName', 'N/A')}</td>
                    <td style="font-size:12px;">{event.get('event_impact_type', 'Informational')}</td>
                </tr>
            """
        
        html_content += """
            </table>
        """
    
    # Add footer with attachment information
    html_content += """
            <div class="summary">
                <h2>Full Report</h2>
                <p>Please see the attached Excel file for complete details on all events.</p>
            </div>
        </div>
    </body>
    </html>
    """
    
    return html_content

def send_ses_email_with_attachment(html_content, excel_buffer, total_events, event_categories, events_analysis, excel_filename=None):
    """
    Send email with Excel attachment using Amazon SES
    
    Args:
        html_content (str): HTML content for email body
        excel_buffer (BytesIO): Excel file as bytes
        total_events (int): Total number of events
        event_categories (dict): Event categories count
        events_analysis (list): List of analyzed events
        excel_filename (str, optional): Custom filename for Excel attachment
        
    Returns:
        None
    """
    try:
        # Get email configuration from environment variables
        sender = os.environ['SENDER_EMAIL']
        recipients_str = os.environ['RECIPIENT_EMAILS']
        recipients = [email.strip() for email in recipients_str.split(',')]
        
        # Create email subject with counts
        critical_count = event_categories.get('critical', 0)
        high_risk_count = sum(1 for event in events_analysis if event.get('risk_level', '').lower() == 'high')
        
        if critical_count > 0:
            subject = f"{customer_name} [CRITICAL] AWS Health Events Analysis - {critical_count} Critical, {high_risk_count} High Risk Events"
        elif high_risk_count > 0:
            subject = f"{customer_name} [HIGH RISK] AWS Health Events Analysis - {high_risk_count} High Risk Events"
        else:
            subject = f"{customer_name} AWS Health Events Analysis - {total_events} Events"
        
        # Generate Excel filename if not provided
        if not excel_filename:
            current_date = datetime.now().strftime('%Y-%m-%d')
            current_time = datetime.now().strftime('%H-%M-%S')
            excel_filename = EXCEL_FILENAME_TEMPLATE.format(date=current_date, time=current_time)
        
        # Create SES client
        ses_client = boto3.client('ses')
        
        # Create message container
        message = {
            'Subject': {
                'Data': subject
            },
            'Body': {
                'Html': {
                    'Data': html_content
                }
            }
        }
        
        # Create raw email message with attachment
        msg_raw = {
            'Source': sender,
            'Destinations': recipients,
            'RawMessage': {
                'Data': create_raw_email_with_attachment(
                    sender=sender,
                    recipients=recipients,
                    subject=subject,
                    html_body=html_content,
                    attachment_data=excel_buffer.getvalue(),
                    attachment_name=excel_filename
                )
            }
        }
        
        # Send email
        response = ses_client.send_raw_email(**msg_raw)
        print(f"Email sent successfully. Message ID: {response['MessageId']}")

        
    except Exception as e:
        print(f"Error sending email: {str(e)}")
        traceback.print_exc()
    
    try:
        # Check if S3 bucket name is configured
        if not S3_BUCKET_NAME:
            return
            
        # Create S3 client
        s3_client = boto3.client('s3')
        
        # Generate S3 key with prefix if provided
        s3_key = f"{S3_KEY_PREFIX.rstrip('/')}/{excel_filename}" if S3_KEY_PREFIX else excel_filename
        
        # Reset buffer position to the beginning
        excel_buffer.seek(0)
        
        # Upload buffer to S3 using put_object
        s3_client.put_object(
            Bucket=S3_BUCKET_NAME,
            Key=s3_key,
            Body=excel_buffer.getvalue(),
            ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
        # Generate S3 URL for the uploaded file
        s3_url = f"s3://{S3_BUCKET_NAME}/{s3_key}"
        
        print(f"Successfully uploaded complete report to {s3_url}")
             
    except Exception as e:
        error_message = f"Error uploading file to S3: {str(e)}"
        print(error_message)


def create_raw_email_with_attachment(sender, recipients, subject, html_body, attachment_data, attachment_name):
    """
    Create raw email with attachment
    
    Args:
        sender (str): Sender email
        recipients (list): List of recipient emails
        subject (str): Email subject
        html_body (str): HTML email body
        attachment_data (bytes): Attachment data
        attachment_name (str): Attachment filename
        
    Returns:
        bytes: Raw email message
    """
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication
    
    # Create message container
    msg = MIMEMultipart('mixed')
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ', '.join(recipients)
    
    # Create HTML part
    msg_body = MIMEMultipart('alternative')
    html_part = MIMEText(html_body, 'html')
    msg_body.attach(html_part)
    msg.attach(msg_body)
    
    # Create attachment part
    att = MIMEApplication(attachment_data)
    att.add_header('Content-Disposition', 'attachment', filename=attachment_name)
    msg.attach(att)
    
    # Convert to string and return
    return msg.as_string().encode('utf-8')

def add_cloudwatch_metrics(event_categories, analyzed_count, total_count, filtered_count):
    """
    Add CloudWatch metrics for monitoring
    
    Args:
        event_categories (dict): Event categories count
        analyzed_count (int): Number of analyzed events
        total_count (int): Total number of events
        filtered_count (int): Number of filtered events
        
    Returns:
        None
    """
    try:
        # Create CloudWatch client
        cloudwatch = boto3.client('cloudwatch')
        
        # Create metrics data
        metrics_data = [
            {
                'MetricName': 'AnalyzedEvents',
                'Value': analyzed_count,
                'Unit': 'Count',
                'Dimensions': [
                    {
                        'Name': 'Function',
                        'Value': 'HealthEventsAnalysis'
                    }
                ]
            },
            {
                'MetricName': 'TotalEvents',
                'Value': total_count,
                'Unit': 'Count',
                'Dimensions': [
                    {
                        'Name': 'Function',
                        'Value': 'HealthEventsAnalysis'
                    }
                ]
            },
            {
                'MetricName': 'FilteredEvents',
                'Value': filtered_count,
                'Unit': 'Count',
                'Dimensions': [
                    {
                        'Name': 'Function',
                        'Value': 'HealthEventsAnalysis'
                    }
                ]
            },
            {
                'MetricName': 'CriticalEvents',
                'Value': event_categories.get('critical', 0),
                'Unit': 'Count',
                'Dimensions': [
                    {
                        'Name': 'Function',
                        'Value': 'HealthEventsAnalysis'
                    }
                ]
            },
            {
                'MetricName': 'HighRiskEvents',
                'Value': event_categories.get('high_risk', 0),
                'Unit': 'Count',
                'Dimensions': [
                    {
                        'Name': 'Function',
                        'Value': 'HealthEventsAnalysis'
                    }
                ]
            }
        ]
        
        # Put metrics data
        cloudwatch.put_metric_data(
            Namespace='AWS/HealthEventsAnalysis',
            MetricData=metrics_data
        )
        
        print("CloudWatch metrics published successfully")
        
    except Exception as e:
        print(f"Error publishing CloudWatch metrics: {str(e)}")
        # Don't raise the exception - metrics are non-critical

def fetch_health_event_details1(event_arn, account_id=None):
    """
    Fetch detailed event information from AWS Health API for any account in the organization
    
    Args:
        event_arn (str): ARN of the health event
        account_id (str, optional): AWS account ID that owns the event
        
    Returns:
        dict: Event details including affected resources
    """
    try:
        health_client = boto3.client('health', region_name='us-east-1')
        
        # First try organization API (works for both current and linked accounts)
        try:
            # Prepare request for organization event details
            org_filter = {
                'eventArn': event_arn
            }
            
            # Add account ID if provided
            if account_id:
                org_filter['awsAccountId'] = account_id
            
            # Get event details using organization API
            org_event_details = health_client.describe_event_details_for_organization(
                organizationEventDetailFilters=[org_filter]
            )
            
            # Get affected entities using organization API
            org_affected_entities = health_client.describe_affected_entities_for_organization(
                organizationEntityFilters=[
                    {
                        'eventArn': event_arn,
                        'awsAccountId': account_id if account_id else get_account_id_from_event(event_arn)
                    }
                ]
            )
            
            # Check if we got successful results
            if org_event_details.get('successfulSet') and len(org_event_details['successfulSet']) > 0:
                return {
                    'details': org_event_details['successfulSet'][0],
                    'entities': org_affected_entities.get('entities', [])
                }
            
            # If we got here, organization API didn't return results
            print(f"Organization API didn't return results for event {event_arn}")
            
        except Exception as org_error:
            print(f"Error using organization API for event {event_arn}: {str(org_error)}")
        
        # Fall back to account-specific API (only works for current account)
        print(f"Falling back to account-specific API for event {event_arn}")
        
        event_details = health_client.describe_event_details(
            eventArns=[event_arn]
        )
        
        affected_entities = health_client.describe_affected_entities(
            filter={
                'eventArns': [event_arn]
            }
        )
        
        return {
            'details': event_details.get('successfulSet', [{}])[0] if event_details.get('successfulSet') else {},
            'entities': affected_entities.get('entities', [])
        }
        
    except Exception as e:
        print(f"Error fetching Health API data: {str(e)}")
        return {'details': {}, 'entities': []}

# Helper function for fetch_health_event_details1
def get_account_id_from_event(event_arn):
    """
    Extract account ID from event ARN if possible
    
    Args:
        event_arn (str): ARN of the health event
        
    Returns:
        str: Account ID or empty string
    """
    try:
        # ARN format: arn:aws:health:region::event/service/id/account-id
        parts = event_arn.split('/')
        if len(parts) >= 4:
            return parts[3]
        return ""
    except Exception:
        return ""

def send_ses_email_with_multiple_attachments(html_content, excel_buffers, excel_filenames, total_events, event_categories, events_analysis):
    """
    Send email with multiple Excel attachments using Amazon SES
    
    Args:
        html_content (str): HTML content for email body
        excel_buffers (dict): Dictionary of BytesIO objects for each account
        excel_filenames (dict): Dictionary of filenames for each account
        total_events (int): Total number of events
        event_categories (dict): Event categories count
        events_analysis (list): List of analyzed events
        
    Returns:
        None
    """
    try:
        # Get email configuration from environment variables
        sender = os.environ['SENDER_EMAIL']
        recipients_str = os.environ['RECIPIENT_EMAILS']
        recipients = [email.strip() for email in recipients_str.split(',')]
        
        # Create email subject with counts
        critical_count = event_categories.get('critical', 0)
        high_risk_count = sum(1 for event in events_analysis if event.get('risk_level', '').lower() == 'high')
        
        if critical_count > 0:
            subject = f"{customer_name} [CRITICAL] AWS Health Events Analysis - {critical_count} Critical, {high_risk_count} High Risk Events"
        elif high_risk_count > 0:
            subject = f"{customer_name} [HIGH RISK] AWS Health Events Analysis - {high_risk_count} High Risk Events"
        else:
            subject = f"{customer_name} AWS Health Events Analysis - {total_events} Events"
        
        # Create SES client
        ses_client = boto3.client('ses')
        
        # Create raw email message with multiple attachments
        msg_raw = {
            'Source': sender,
            'Destinations': recipients,
            'RawMessage': {
                'Data': create_raw_email_with_multiple_attachments(
                    sender=sender,
                    recipients=recipients,
                    subject=subject,
                    html_body=html_content,
                    excel_buffers=excel_buffers,
                    excel_filenames=excel_filenames
                )
            }
        }
        
        # Send email
        response = ses_client.send_raw_email(**msg_raw)
        print(f"Email with multiple attachments sent successfully. Message ID: {response['MessageId']}")

        # Upload files to S3
        if S3_BUCKET_NAME:
            s3_client = boto3.client('s3')
            
            for account_id, excel_buffer in excel_buffers.items():
                filename = excel_filenames[account_id]
                
                # Generate S3 key with prefix if provided
                s3_key = f"{S3_KEY_PREFIX.rstrip('/')}/{filename}" if S3_KEY_PREFIX else filename
                
                # Reset buffer position to the beginning
                excel_buffer.seek(0)
                
                # Upload buffer to S3 using put_object
                s3_client.put_object(
                    Bucket=S3_BUCKET_NAME,
                    Key=s3_key,
                    Body=excel_buffer.getvalue(),
                    ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                
                # Generate S3 URL for the uploaded file
                s3_url = f"s3://{S3_BUCKET_NAME}/{s3_key}"
                print(f"Successfully uploaded file to {s3_url}")
                
    except Exception as e:
        print(f"Error sending email with multiple attachments: {str(e)}")
        traceback.print_exc()

def create_raw_email_with_multiple_attachments(sender, recipients, subject, html_body, excel_buffers, excel_filenames):
    """
    Create raw email with multiple attachments
    
    Args:
        sender (str): Sender email
        recipients (list): List of recipient emails
        subject (str): Email subject
        html_body (str): HTML email body
        excel_buffers (dict): Dictionary of BytesIO objects for each account
        excel_filenames (dict): Dictionary of filenames for each account
        
    Returns:
        bytes: Raw email message
    """
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication
    
    # Create message container
    msg = MIMEMultipart('mixed')
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ', '.join(recipients)
    
    # Create HTML part
    msg_body = MIMEMultipart('alternative')
    html_part = MIMEText(html_body, 'html')
    msg_body.attach(html_part)
    msg.attach(msg_body)
    
    # Create attachment parts for each Excel file
    for account_id, excel_buffer in excel_buffers.items():
        filename = excel_filenames[account_id]
        att = MIMEApplication(excel_buffer.getvalue())
        att.add_header('Content-Disposition', 'attachment', filename=filename)
        msg.attach(att)
    
    # Convert to string and return
    return msg.as_string().encode('utf-8')

def parse_group_config(config_str):
    """
    Parse group configuration from environment variable string
    
    Args:
        config_str (str): Configuration string in format "Group Name 1 = value1, value2; Group Name 2 = value3, value4"
        
    Returns:
        dict: Dictionary mapping group names to lists of values
    """
    if not config_str:
        return {}
    
    result = {}
    # Split by semicolon to get each group definition
    group_defs = config_str.split(';')
    
    for group_def in group_defs:
        group_def = group_def.strip()
        if not group_def or '=' not in group_def:
            continue
            
        # Split by equals sign to get group name and values
        parts = group_def.split('=', 1)
        if len(parts) != 2:
            continue
            
        group_name = parts[0].strip()
        values_str = parts[1].strip()
        
        # Split values by comma
        values = [v.strip() for v in values_str.split(',') if v.strip()]
        
        if group_name and values:
            result[group_name] = values
    
    return result

def validate_group_configs(account_groups, email_groups):
    """
    Validate that group configurations are synchronized
    
    Args:
        account_groups (dict): Dictionary mapping group names to account IDs
        email_groups (dict): Dictionary mapping group names to email addresses
        
    Returns:
        bool: True if configurations are valid and synchronized, False otherwise
    """
    # Check if both configs are present
    if not account_groups or not email_groups:
        return False
    
    # Check if group names match
    account_group_names = set(account_groups.keys())
    email_group_names = set(email_groups.keys())
    
    return account_group_names == email_group_names

def group_events_by_account_groups(events_analysis, account_groups):
    """
    Group events by account groups
    
    Args:
        events_analysis (list): List of analyzed events
        account_groups (dict): Dictionary mapping group names to account IDs
        
    Returns:
        dict: Dictionary mapping group names to lists of events
    """
    # Create a mapping from account ID to group name for faster lookups
    account_to_group = {}
    for group_name, account_ids in account_groups.items():
        for account_id in account_ids:
            account_to_group[account_id] = group_name
    
    # Create a dictionary to hold events for each group
    group_events = {group_name: [] for group_name in account_groups.keys()}
    # Add a key for "other" accounts not in any group
    group_events['other'] = []
    
    # Distribute events to their respective groups
    for event in events_analysis:
        account_id = event.get('accountId', 'N/A')
        if account_id in account_to_group:
            group_name = account_to_group[account_id]
            group_events[group_name].append(event)
        else:
            group_events['other'].append(event)
    
    return group_events

def generate_group_summary_html(group_events, total_events, filtered_events, category_filter):
    """
    Generate HTML summary for email specific to a group
    
    Args:
        group_events (list): List of events for this group
        total_events (int): Total number of events (across all groups)
        filtered_events (int): Number of filtered events
        category_filter (list): Categories used for filtering
        
    Returns:
        str: HTML content for email
    """
    # Current date for the report
    current_date = datetime.now().strftime('%Y-%m-%d')
    
    # Calculate event counts for this group
    critical_count = sum(1 for event in group_events if event.get('critical', False))
    high_risk_count = sum(1 for event in group_events if event.get('risk_level', '').lower() == 'high')
    medium_risk_count = sum(1 for event in group_events if event.get('risk_level', '').lower() == 'medium')
    low_risk_count = sum(1 for event in group_events if event.get('risk_level', '').lower() == 'low')

    # Define CSS for better table formatting
    table_css = """
    <style>
        .health-events-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 5px;
        }
        .health-events-table th, .health-events-table td {
            border: 1px solid #ddd;
            padding: 6px;
            text-align: left;
            vertical-align: top;
            word-wrap: break-word;
            font-size: 5px;
        }
        .health-events-table th {
            background-color: #f2f2f2;
        }
        /* Column width constraints */
        .col-arn {
            width: 15%;
        }
        .col-region {
            width: 8%;
        }
        .col-start-time {
            width: 12%;
        }
        .col-risk {
            width: 8%;
        }
        .col-accountid {
            width: 12%;
        }
    </style>
    """
    
    # Start building HTML content
    html_content = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; }}
            .header {{ background-color: #232F3E; color: white; padding: 20px; }}
            .content {{ padding: 20px; }}
            table {{ border-collapse: collapse; width: 100%; margin-top: 20px; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left;"font-size:5px;" }}
            th {{ background-color: #f2f2f2; }}
            .critical {{ background-color: #ffcccc; }}
            .high {{ background-color: #fff2cc; }}
            .medium {{ background-color: #e6f2ff; }}
            .summary {{ margin-top: 20px; margin-bottom: 20px; }}
        </style>
    </head>
    <body>
        <div class="header">
            <h1>AWS Health Events Analysis Report</h1>
            <p>Date: {current_date}</p>
        </div>
        <div class="content">
            <div class="summary">
                <h2>Summary</h2>
                <p>Total AWS Health events analyzed for your accounts: {len(group_events)} (out of {total_events} total events found)</p>
    """
    
    # Add analysis window information
    end_time = datetime.utcnow()
    start_time = end_time - timedelta(days=int(os.environ['ANALYSIS_WINDOW_DAYS']))
    html_content += f"""
                <p>Analysis Window: {start_time.strftime('%Y-%m-%d %H:%M:%S')} UTC to {end_time.strftime('%Y-%m-%d %H:%M:%S')} UTC</p>
    """
    
    # Add filter information if applicable
    if category_filter:
        html_content += f"""
                <p>Events filtered by categories: {', '.join(category_filter)}</p>
                <p>Events excluded by filter: {filtered_events}</p>
        """
    
    # Add event category counts
    html_content += """
                <h3>Event Categories</h3>
                <ul>
    """
    
    # Add critical events count if any
    if critical_count > 0:
        html_content += f"""
                <li><strong>Critical Events:</strong> {critical_count}</li>
        """
    
    # Add risk level counts
    html_content += f"""
                <li><strong>High Risk Events:</strong> {high_risk_count}</li>
                <li><strong>Medium Risk Events:</strong> {medium_risk_count}</li>
                <li><strong>Low Risk Events:</strong> {low_risk_count}</li>
            </ul>
            </div>
    """
    
    # Add critical events table if any exist
    critical_events = [event for event in group_events if event.get('critical', False)]
    if critical_events:
        html_content += """
            <h2>Critical Events</h2>
            <table class="health-events-table">
                <tr>
                    <th style="font-size:12px;">Event ARN</th>
                    <th style="font-size:12px;">Region</th>
                    <th style="font-size:12px;">Start Time</th>
                    <th style="font-size:12px;">Risk Level</th>
                    <th style="font-size:12px;">Account ID</th>
                    <th style="font-size:12px;">Account Name</th>
                    <th style="font-size:12px;">Event Impact Type</th>
                </tr>
        """
        
        for event in critical_events:
            # Get event ARN, preferring eventArn if available, falling back to arn
            event_arn = event.get('eventArn', event.get('arn', 'N/A'))
            
            html_content += f"""
                <tr class="critical">
                    <td style="font-size:12px;">{event_arn}</td>
                    <td style="font-size:12px;">{event.get('region', 'N/A')}</td>
                    <td style="font-size:12px;">{event.get('start_time', 'N/A')}</td>
                    <td style="font-size:12px;">{event.get('risk_level', 'N/A').upper()}</td>
                    <td style="font-size:12px;">{event.get('accountId', 'N/A')}</td>
                    <td style="font-size:12px;">{event.get('accountName', 'N/A')}</td>
                    <td style="font-size:12px;">{event.get('event_impact_type', 'Informational')}</td>
                </tr>
            """
        
        html_content += """
            </table>
        """
    
    # Add high risk events table
    high_risk_events = [event for event in group_events if event.get('risk_level', '').lower() == 'high']
    if high_risk_events:
        html_content += """
            <h2>High Risk Events</h2>
            <table class="health-events-table">
                <tr>
                    <th style="font-size:12px;">Event ARN</th>
                    <th style="font-size:12px;">Region</th>
                    <th style="font-size:12px;">Start Time</th>
                    <th style="font-size:12px;">Risk Level</th>
                    <th style="font-size:12px;">Account ID</th>
                    <th style="font-size:12px;">Account Name</th>
                    <th style="font-size:12px;">Event Impact Type</th>
                </tr>
        """
        
        for event in high_risk_events:
            # Get event ARN, preferring eventArn if available, falling back to arn
            event_arn = event.get('eventArn', event.get('arn', 'N/A'))
            
            html_content += f"""
                <tr class="high">
                    <td style="font-size:12px;">{event_arn}</td>
                    <td style="font-size:12px;">{event.get('region', 'N/A')}</td>
                    <td style="font-size:12px;">{event.get('start_time', 'N/A')}</td>
                    <td style="font-size:12px;">{event.get('risk_level', 'N/A').upper()}</td>
                    <td style="font-size:12px;">{event.get('accountId', 'N/A')}</td>
                    <td style="font-size:12px;">{event.get('accountName', 'N/A')}</td>
                    <td style="font-size:12px;">{event.get('event_impact_type', 'Informational')}</td>
                </tr>
            """
        
        html_content += """
            </table>
        """
    
    # Add footer with attachment information
    html_content += """
            <div class="summary">
                <h2>Full Report</h2>
                <p>Please see the attached Excel file for complete details on all events.</p>
            </div>
        </div>
    </body>
    </html>
    """
    
    return html_content

def send_ses_email_with_group_attachment(html_content, excel_buffer, excel_filename, total_events, filtered_events, category_filter, group_events, group_recipients):
    """
    Send email with Excel attachment to group-specific recipients using Amazon SES
    
    Args:
        html_content (str): HTML content for email body
        excel_buffer (BytesIO): Excel file as bytes
        excel_filename (str): Filename for the Excel attachment
        total_events (int): Total number of events across all groups
        filtered_events (int): Number of filtered events
        category_filter (list): Categories used for filtering
        group_events (list): List of events for this specific group
        group_recipients (list): List of email recipients for this group
        
    Returns:
        None
    """
    try:
        # Get email configuration from environment variables
        sender = os.environ['SENDER_EMAIL']
        
        # Generate group-specific HTML content
        group_html = generate_group_summary_html(group_events, total_events, filtered_events, category_filter)
        
        # Create email subject with counts
        critical_count = sum(1 for event in group_events if event.get('critical', False))
        high_risk_count = sum(1 for event in group_events if event.get('risk_level', '').lower() == 'high')
        
        if critical_count > 0:
            subject = f"{customer_name} [CRITICAL] AWS Health Events Analysis - {critical_count} Critical, {high_risk_count} High Risk Events"
        elif high_risk_count > 0:
            subject = f"{customer_name} [HIGH RISK] AWS Health Events Analysis - {high_risk_count} High Risk Events"
        else:
            subject = f"{customer_name} AWS Health Events Analysis - {len(group_events)} Events"
        
        # Create SES client
        ses_client = boto3.client('ses')
        
        # Create raw email message with attachment
        msg_raw = {
            'Source': sender,
            'Destinations': group_recipients,
            'RawMessage': {
                'Data': create_raw_email_with_attachment(
                    sender=sender,
                    recipients=group_recipients,
                    subject=subject,
                    html_body=group_html,
                    attachment_data=excel_buffer.getvalue(),
                    attachment_name=excel_filename
                )
            }
        }
        
        # Send email
        response = ses_client.send_raw_email(**msg_raw)
        print(f"Email sent successfully to {', '.join(group_recipients)}. Message ID: {response['MessageId']}")
        
    except Exception as e:
        print(f"Error sending email to group: {str(e)}")
        traceback.print_exc()

# Add this function to handle DynamoDB storage
def process_single_event(bedrock_client, event_data):
    """
    Process a single event for analysis and DynamoDB storage
    
    Args:
        bedrock_client: Amazon Bedrock client
        event_data (dict): Event data to process
        
    Returns:
        list: List containing the analyzed event data or empty list if processing failed
    """
    try:
        print(f"Processing single event: {event_data.get('arn', 'unknown')}")
        
        # Get account ID and name
        account_id = event_data.get('accountId', 'N/A')
        account_name = get_account_name(account_id) if account_id != 'N/A' else 'N/A'
        
        # Fetch additional details from Health API if needed
        if account_id != 'N/A':
            health_data = fetch_health_event_details1(event_data.get('arn', ''), account_id)
            
            # Extract affected resources
            affected_resources = extract_affected_resources(health_data.get('entities', []))
            
            # Check if we have a better description from the health data
            health_description = health_data.get('details', {}).get('eventDescription', {}).get('latestDescription', '')
            if health_description:
                event_data['description'] = health_description
        else:
            affected_resources = 'None specified'
        
        # Make sure we have a description
        if not event_data.get('description'):
            event_data['description'] = 'No description available'
        
        # Analyze the event with Bedrock
        analysis = analyze_event_with_bedrock(bedrock_client, event_data)
        
        # Categorize the analysis
        categories = categorize_analysis(analysis)
        
        # Create structured event data
        event_entry = {
            "arn": event_data.get('arn', 'N/A'),
            "eventArn": event_data.get('arn', 'N/A'),
            "event_type": event_data.get('eventTypeCode', 'N/A'),
            "description": event_data.get('description', 'N/A'),
            "region": event_data.get('region', 'N/A'),
            "start_time": format_time(event_data.get('startTime', 'N/A')),
            "last_update_time": format_time(event_data.get('lastUpdatedTime', 'N/A')),
            "event_type_category": event_data.get('eventTypeCategory', 'N/A'),
            "service": event_data.get('service', 'N/A'),
            "analysis_text": analysis,
            "critical": categories.get('critical', False),
            "risk_level": categories.get('risk_level', 'low'),
            "accountId": account_id,
            "accountName": account_name,
            "impact_analysis": categories.get('impact_analysis', ''),
            "required_actions": categories.get('required_actions', ''),
            "time_sensitivity": categories.get('time_sensitivity', 'Routine'),
            "risk_category": categories.get('risk_category', 'Unknown'),
            "consequences_if_ignored": categories.get('consequences_if_ignored', ''),
            "affected_resources": affected_resources,
            "event_impact_type": categories.get('event_impact_type', 'Unknown')
        }
        
        print(f"Successfully analyzed single event: {event_data.get('arn', 'unknown')}")
        return [event_entry]
        
    except Exception as e:
        print(f"Error processing single event: {str(e)}")
        traceback.print_exc()
        return []

def store_events_in_dynamodb(events_analysis):
    """
    Store analyzed events in DynamoDB table
    
    Args:
        events_analysis (list): List of analyzed events data
        
    Returns:
        dict: Summary of storage operation
    """
    if not DYNAMODB_TABLE_NAME:
        print("DynamoDB table name not provided, skipping storage")
        return {"stored": 0, "failed": 0, "updated": 0}
    
    print(f"Storing {len(events_analysis)} events in DynamoDB table: {DYNAMODB_TABLE_NAME}")
    
    # Create DynamoDB resource
    dynamodb = boto3.resource('dynamodb')
    table = dynamodb.Table(DYNAMODB_TABLE_NAME)
    
    # Track success and failures
    stored_count = 0
    failed_count = 0
    updated_count = 0
    
    # Get current timestamp for metadata
    analysis_timestamp = datetime.utcnow().isoformat()
    
    # Process each event
    for event in events_analysis:
        try:
            # Get primary key values
            event_arn = event.get('eventArn', event.get('arn', ''))
            account_id = event.get('accountId', 'N/A')
            
            if not event_arn:
                print("Skipping event with no ARN")
                failed_count += 1
                continue
            
            # Create item with all relevant fields
            item = {
                'eventArn': event_arn,
                'accountId': account_id,
                'eventType': event.get('event_type', 'N/A'),
                'eventTypeCategory': event.get('event_type_category', 'N/A'),
                'region': event.get('region', 'N/A'),
                'service': event.get('service', 'N/A'),
                'startTime': event.get('start_time', 'N/A'),
                'lastUpdateTime': event.get('last_update_time', 'N/A'),
                'description': event.get('description', 'N/A'),
                'critical': event.get('critical', False),
                'riskLevel': event.get('risk_level', 'low'),
                'accountName': event.get('accountName', 'N/A'),
                'timeSensitivity': event.get('time_sensitivity', 'Routine'),
                'riskCategory': event.get('risk_category', 'Unknown'),
                'eventImpactType': event.get('event_impact_type', 'Informational'),
                'requiredActions': event.get('required_actions', ''),
                'impactAnalysis': event.get('impact_analysis', ''),
                'consequencesIfIgnored': event.get('consequences_if_ignored', ''),
                'affectedResources': event.get('affected_resources', 'None specified'),
                'analysisTimestamp': analysis_timestamp,
                'analysisVersion': '1.0',
                'Category': None  # New column, initially set to None
            }
            
            # Convert any empty strings to None (null in DynamoDB)
            for key, value in item.items():
                if value == '':
                    item[key] = None
            
            # Handle decimal conversion for numeric values
            item = json.loads(json.dumps(item), parse_float=Decimal)
            
            # Check if the item already exists
            try:
                response = table.get_item(
                    Key={
                        'eventArn': event_arn,
                        'accountId': account_id
                    }
                )
                
                if 'Item' in response:
                    print(f"Event {event_arn} for account {account_id} already exists, updating...")
                    # Update the existing item
                    table.put_item(Item=item)
                    updated_count += 1
                else:
                    # Store new item in DynamoDB
                    table.put_item(Item=item)
                    stored_count += 1
            except Exception as e:
                print(f"Error checking for existing item: {str(e)}")
                # Fall back to put_item
                table.put_item(Item=item)
                stored_count += 1
            
        except Exception as e:
            print(f"Error storing event in DynamoDB: {str(e)}")
            traceback.print_exc()
            failed_count += 1
    
    print(f"DynamoDB storage complete: {stored_count} stored, {updated_count} updated, {failed_count} failed")
    return {"stored": stored_count, "updated": updated_count, "failed": failed_count}

