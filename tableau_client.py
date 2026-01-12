import requests
import xml.etree.ElementTree as ET
from typing import List, Dict
import csv


class TableauCloudClient:
    """Tableau REST API Client (supports both Cloud and Server)"""
    
    def __init__(self, pod_name: str, api_version: str = "3.23"):
        if pod_name.startswith('http'):
            self.server_url = pod_name.rstrip('/')
        else:
            self.server_url = f"https://{pod_name}.online.tableau.com"
        
        self.api_version = api_version
        self.base_url = f"{self.server_url}/api/{api_version}"
        self.auth_token = None
        self.site_id = None
        self.user_id = None
    
    def sign_in(self, pat_name: str, pat_secret: str, site_name: str = ""):
        """Sign in to Tableau with PAT (Personal Access Token)"""
        url = f"{self.base_url}/auth/signin"
        
        payload = f"""<tsRequest>
    <credentials personalAccessTokenName="{pat_name}" personalAccessTokenSecret="{pat_secret}">
        <site contentUrl="{site_name}" />
    </credentials>
</tsRequest>"""
        
        headers = {'Content-Type': 'application/xml'}
        response = requests.post(url, data=payload, headers=headers)
        response.raise_for_status()
        
        root = ET.fromstring(response.content)
        ns = {'t': 'http://tableau.com/api'}
        
        error = root.find('.//t:error', ns)
        if error is not None:
            error_summary = error.find('t:summary', ns)
            error_msg = error_summary.text if error_summary is not None else "Authentication error"
            raise Exception(error_msg)
        
        credentials = root.find('.//t:credentials', ns)
        site = root.find('.//t:site', ns)
        user = root.find('.//t:user', ns)
        
        if credentials is not None and site is not None:
            self.auth_token = credentials.get('token')
            self.site_id = site.get('id')
            if user is not None:
                self.user_id = user.get('id')
            return True
        else:
            raise Exception("Failed to retrieve authentication information")
    
    def _get_headers(self) -> Dict[str, str]:
        """Get authentication headers"""
        if not self.auth_token:
            raise Exception("Please execute sign_in() first")
        return {
            'X-Tableau-Auth': self.auth_token,
            'Accept': 'application/xml'
        }
    
    def _parse_permissions(self, xml_content: bytes) -> List[Dict]:
        """Parse Permission Rules XML"""
        root = ET.fromstring(xml_content)
        ns = {'t': 'http://tableau.com/api'}
        permissions = []
        
        for grantee in root.findall('.//t:granteeCapabilities', ns):
            perm_entry = {'capabilities': []}
            
            user = grantee.find('t:user', ns)
            group = grantee.find('t:group', ns)
            
            if user is not None:
                perm_entry['grantee_type'] = 'user'
                perm_entry['grantee_id'] = user.get('id')
            elif group is not None:
                perm_entry['grantee_type'] = 'group'
                perm_entry['grantee_id'] = group.get('id')
            
            capabilities = grantee.find('t:capabilities', ns)
            if capabilities is not None:
                for capability in capabilities.findall('t:capability', ns):
                    perm_entry['capabilities'].append({
                        'name': capability.get('name'),
                        'mode': capability.get('mode')
                    })
            
            permissions.append(perm_entry)
        
        return permissions
    
    def get_projects(self) -> List[Dict]:
        """Get list of projects"""
        url = f"{self.base_url}/sites/{self.site_id}/projects"
        response = requests.get(url, headers=self._get_headers())
        response.raise_for_status()
        
        root = ET.fromstring(response.content)
        ns = {'t': 'http://tableau.com/api'}
        projects = []
        
        for project in root.findall('.//t:project', ns):
            proj_data = {
                'id': project.get('id'),
                'name': project.get('name'),
                'parentProjectId': project.get('parentProjectId'),
                'contentPermissions': project.get('contentPermissions'),
            }
            projects.append(proj_data)
        
        return projects
    
    def get_workbooks(self, project_id: str = None) -> List[Dict]:
        """Get list of workbooks"""
        url = f"{self.base_url}/sites/{self.site_id}/workbooks"
        response = requests.get(url, headers=self._get_headers())
        response.raise_for_status()
        
        root = ET.fromstring(response.content)
        ns = {'t': 'http://tableau.com/api'}
        workbooks = []
        
        for workbook in root.findall('.//t:workbook', ns):
            project = workbook.find('t:project', ns)
            proj_id = project.get('id') if project is not None else None
            
            # Filter by project_id if specified
            if project_id is None or proj_id == project_id:
                wb_data = {
                    'id': workbook.get('id'),
                    'name': workbook.get('name'),
                    'project_id': proj_id,
                    'project_name': project.get('name') if project is not None else ''
                }
                workbooks.append(wb_data)
        
        return workbooks
    
    def get_workbook_by_id(self, workbook_id: str) -> Dict:
        """Get detailed information for a specific workbook"""
        url = f"{self.base_url}/sites/{self.site_id}/workbooks/{workbook_id}"
        response = requests.get(url, headers=self._get_headers())
        response.raise_for_status()
        
        root = ET.fromstring(response.content)
        ns = {'t': 'http://tableau.com/api'}
        workbook = root.find('.//t:workbook', ns)
        
        if workbook is None:
            return None
        
        wb_data = {
            'id': workbook.get('id'),
            'name': workbook.get('name'),
        }
        
        views = workbook.findall('.//t:view', ns)
        if views:
            wb_data['views'] = [{
                'id': view.get('id'),
                'name': view.get('name'),
            } for view in views]
        
        return wb_data
    
    def get_datasources(self, project_id: str = None) -> List[Dict]:
        """Get list of data sources"""
        url = f"{self.base_url}/sites/{self.site_id}/datasources"
        response = requests.get(url, headers=self._get_headers())
        response.raise_for_status()
        
        root = ET.fromstring(response.content)
        ns = {'t': 'http://tableau.com/api'}
        datasources = []
        
        for datasource in root.findall('.//t:datasource', ns):
            project = datasource.find('t:project', ns)
            proj_id = project.get('id') if project is not None else None
            
            if project_id is None or proj_id == project_id:
                ds_data = {
                    'id': datasource.get('id'),
                    'name': datasource.get('name'),
                    'project_id': proj_id,
                    'project_name': project.get('name') if project is not None else ''
                }
                datasources.append(ds_data)
        
        return datasources
    
    def get_flows(self, project_id: str = None) -> List[Dict]:
        """Get list of flows"""
        url = f"{self.base_url}/sites/{self.site_id}/flows"
        response = requests.get(url, headers=self._get_headers())
        response.raise_for_status()
        
        root = ET.fromstring(response.content)
        ns = {'t': 'http://tableau.com/api'}
        flows = []
        
        for flow in root.findall('.//t:flow', ns):
            project = flow.find('t:project', ns)
            proj_id = project.get('id') if project is not None else None
            
            if project_id is None or proj_id == project_id:
                flow_data = {
                    'id': flow.get('id'),
                    'name': flow.get('name'),
                    'project_id': proj_id,
                    'project_name': project.get('name') if project is not None else ''
                }
                flows.append(flow_data)
        
        return flows
    
    def get_groups(self) -> Dict[str, str]:
        """Get list of groups and return ID->name mapping"""
        url = f"{self.base_url}/sites/{self.site_id}/groups"
        response = requests.get(url, headers=self._get_headers())
        response.raise_for_status()
        
        root = ET.fromstring(response.content)
        ns = {'t': 'http://tableau.com/api'}
        
        groups = {}
        for group in root.findall('.//t:group', ns):
            group_id = group.get('id')
            group_name = group.get('name')
            if group_id and group_name:
                groups[group_id] = group_name
        
        return groups
    
    def get_group_users(self, group_id: str) -> List[Dict]:
        """Get users in a specific group"""
        url = f"{self.base_url}/sites/{self.site_id}/groups/{group_id}/users"
        response = requests.get(url, headers=self._get_headers())
        response.raise_for_status()
        
        root = ET.fromstring(response.content)
        ns = {'t': 'http://tableau.com/api'}
        
        users = []
        for user in root.findall('.//t:user', ns):
            user_data = {
                'id': user.get('id'),
                'name': user.get('name'),
                'email': user.get('email', '')
            }
            users.append(user_data)
        
        return users
    
    def export_group_members(self, folder_path: str, timestamp: str) -> str:
        """Export group membership information to CSV"""
        try:
            # Get all groups
            groups = self.get_groups()
            
            all_members = []
            
            for group_id, group_name in groups.items():
                try:
                    # Get users in this group
                    members = self.get_group_users(group_id)
                    
                    for member in members:
                        all_members.append({
                            'group_id': group_id,
                            'group_name': group_name,
                            'user_id': member['id'],
                            'user_name': member.get('name', ''),
                            'user_email': member.get('email', ''),
                        })
                except Exception as e:
                    print(f"Warning: Failed to get members for group {group_name}: {e}")
            
            # Export to CSV
            if all_members:
                filename = f"{folder_path}/group_members_{timestamp}.csv"
                
                with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
                    fieldnames = ['group_id', 'group_name', 'user_id', 'user_name', 'user_email']
                    writer = csv.DictWriter(f, fieldnames=fieldnames)
                    writer.writeheader()
                    writer.writerows(all_members)
                
                print(f"Exported {len(all_members)} group memberships")
                return filename
            
            return None
        except Exception as e:
            print(f"Error exporting group members: {e}")
            return None
    
    def get_users(self) -> Dict[str, str]:
        """Get list of users and return ID->name mapping"""
        url = f"{self.base_url}/sites/{self.site_id}/users"
        response = requests.get(url, headers=self._get_headers())
        response.raise_for_status()
        
        root = ET.fromstring(response.content)
        ns = {'t': 'http://tableau.com/api'}
        
        users = {}
        for user in root.findall('.//t:user', ns):
            user_id = user.get('id')
            user_name = user.get('name')
            if user_id and user_name:
                users[user_id] = user_name
        
        return users
    
    def get_workbook_permissions(self, workbook_id: str) -> List[Dict]:
        """Get Permission Rules for a workbook"""
        url = f"{self.base_url}/sites/{self.site_id}/workbooks/{workbook_id}/permissions"
        response = requests.get(url, headers=self._get_headers())
        response.raise_for_status()
        return self._parse_permissions(response.content)
    
    def get_datasource_permissions(self, datasource_id: str) -> List[Dict]:
        """Get Permission Rules for a data source"""
        url = f"{self.base_url}/sites/{self.site_id}/datasources/{datasource_id}/permissions"
        response = requests.get(url, headers=self._get_headers())
        response.raise_for_status()
        return self._parse_permissions(response.content)
    
    def get_view_permissions(self, view_id: str) -> List[Dict]:
        """Get Permission Rules for a view"""
        url = f"{self.base_url}/sites/{self.site_id}/views/{view_id}/permissions"
        response = requests.get(url, headers=self._get_headers())
        response.raise_for_status()
        return self._parse_permissions(response.content)
    
    def get_flow_permissions(self, flow_id: str) -> List[Dict]:
        """Get Permission Rules for a flow"""
        url = f"{self.base_url}/sites/{self.site_id}/flows/{flow_id}/permissions"
        response = requests.get(url, headers=self._get_headers())
        response.raise_for_status()
        return self._parse_permissions(response.content)
    
    def get_project_permissions(self, project_id: str) -> List[Dict]:
        """Get Permission Rules for a project including default permissions for all content types"""
        # Define permission templates
        project_templates = {
            'InheritedProjectLeader': {
                # Administer template - all permissions
                'Read': 'Allow', 'Write': 'Allow',
            },
        }
        
        workbook_templates = {
            'InheritedProjectLeader': {
                # Administer template for workbooks
                'workbook_Read': 'Allow', 'workbook_Filter': 'Allow', 'workbook_ViewComments': 'Allow',
                'workbook_AddComment': 'Allow', 'workbook_ExportImage': 'Allow', 'workbook_ExportData': 'Allow',
                'workbook_ShareView': 'Allow', 'workbook_ViewUnderlyingData': 'Allow', 'workbook_WebAuthoring': 'Allow',
                'workbook_RunExplainData': 'Allow', 'workbook_ExportXml': 'Allow', 'workbook_Write': 'Allow',
                'workbook_ChangeHierarchy': 'Allow', 'workbook_Delete': 'Allow', 'workbook_ChangePermissions': 'Allow',
                'workbook_ExtractRefresh': 'Allow',
            },
        }
        
        datasource_templates = {
            'InheritedProjectLeader': {
                # Administer template for datasources
                'datasource_Read': 'Allow', 'datasource_Connect': 'Allow', 'datasource_ExportXml': 'Allow',
                'datasource_Write': 'Allow', 'datasource_SaveAs': 'Allow', 'datasource_VizqlDataApiAccess': 'Allow',
                'datasource_PulseMetricDefine': 'Allow', 'datasource_ChangeHierarchy': 'Allow',
                'datasource_Delete': 'Allow', 'datasource_ChangePermissions': 'Allow', 'datasource_ExtractRefresh': 'Allow',
            },
        }
        
        flow_templates = {
            'InheritedProjectLeader': {
                # Administer template for flows
                'flow_Read': 'Allow', 'flow_ExportXml': 'Allow', 'flow_Execute': 'Allow',
                'flow_Write': 'Allow', 'flow_WebAuthoringForFlows': 'Allow', 'flow_ChangeHierarchy': 'Allow',
                'flow_Delete': 'Allow', 'flow_ChangePermissions': 'Allow',
            },
        }
        
        virtualconnection_templates = {
            'InheritedProjectLeader': {
                # Administer template for virtual connections
                # Note: Write capability has no API endpoint, marked as special value
                'virtualconnection_Read': 'Allow', 'virtualconnection_Connect': 'Allow', 
                'virtualconnection_Write': 'No API Endpoint', 'virtualconnection_ChangeHierarchy': 'Allow',
                'virtualconnection_Delete': 'Allow', 'virtualconnection_ChangePermissions': 'Allow',
            },
        }
        
        database_templates = {
            'InheritedProjectLeader': {
                # Administer template for databases
                'database_Read': 'Allow', 'database_Write': 'Allow', 'database_ChangeHierarchy': 'Allow',
                'database_ChangePermissions': 'Allow',
            },
        }
        
        table_templates = {
            'InheritedProjectLeader': {
                # Administer template for tables
                'table_Read': 'Allow', 'table_Write': 'Allow', 'table_ChangeHierarchy': 'Allow',
                'table_ChangePermissions': 'Allow',
            },
        }
        
        # Get project info for contentPermissions
        projects = self.get_projects()
        proj = next((p for p in projects if p['id'] == project_id), None)
        content_permissions = proj.get('contentPermissions', '') if proj else ''
        
        # Get project-level permissions (Projects Tab)
        url = f"{self.base_url}/sites/{self.site_id}/projects/{project_id}/permissions"
        response = requests.get(url, headers=self._get_headers())
        response.raise_for_status()
        project_perms = self._parse_permissions(response.content)
        
        # Get default permissions for different content types
        workbook_perms = self._get_default_permissions(project_id, 'workbooks')
        datasource_perms = self._get_default_permissions(project_id, 'datasources')
        flow_perms = self._get_default_permissions(project_id, 'flows')
        
        # Try multiple endpoints for virtual connections
        virtualconnection_perms = []
        for vc_endpoint in ['virtualconnections', 'lenses', 'virtual-connections']:
            virtualconnection_perms = self._get_default_permissions(project_id, vc_endpoint)
            if virtualconnection_perms:
                break
        
        database_perms = self._get_default_permissions(project_id, 'databases')
        table_perms = self._get_default_permissions(project_id, 'tables')
        
        # Merge all permissions by grantee - use dict for capabilities
        grantee_permissions = {}
        grantees_with_template = set()  # Track which grantees have template applied
        grantees_with_virtualconnection = set()  # Track which grantees have any VC permissions
        
        # Project permissions
        for perm in project_perms:
            grantee_key = (perm['grantee_type'], perm['grantee_id'])
            if grantee_key not in grantee_permissions:
                grantee_permissions[grantee_key] = {
                    'grantee_type': perm['grantee_type'],
                    'grantee_id': perm['grantee_id'],
                    'capabilities': {}
                }
            for cap in perm['capabilities']:
                cap_name = cap['name']
                cap_mode = cap['mode']
                
                # Check if this is a template capability
                if cap_name in project_templates:
                    # Mark this grantee as having template
                    grantees_with_template.add(grantee_key)
                    # Expand template to individual capabilities
                    for expanded_cap, expanded_mode in project_templates[cap_name].items():
                        grantee_permissions[grantee_key]['capabilities'][expanded_cap] = expanded_mode
                    # Also expand for all other tabs
                    for expanded_cap, expanded_mode in workbook_templates[cap_name].items():
                        grantee_permissions[grantee_key]['capabilities'][expanded_cap] = expanded_mode
                    for expanded_cap, expanded_mode in datasource_templates[cap_name].items():
                        grantee_permissions[grantee_key]['capabilities'][expanded_cap] = expanded_mode
                    for expanded_cap, expanded_mode in flow_templates[cap_name].items():
                        grantee_permissions[grantee_key]['capabilities'][expanded_cap] = expanded_mode
                    for expanded_cap, expanded_mode in virtualconnection_templates[cap_name].items():
                        grantee_permissions[grantee_key]['capabilities'][expanded_cap] = expanded_mode
                    for expanded_cap, expanded_mode in database_templates[cap_name].items():
                        grantee_permissions[grantee_key]['capabilities'][expanded_cap] = expanded_mode
                    for expanded_cap, expanded_mode in table_templates[cap_name].items():
                        grantee_permissions[grantee_key]['capabilities'][expanded_cap] = expanded_mode
                else:
                    grantee_permissions[grantee_key]['capabilities'][cap_name] = cap_mode
        
        # Workbook default permissions
        for perm in workbook_perms:
            grantee_key = (perm['grantee_type'], perm['grantee_id'])
            
            # Skip if this grantee already has template applied
            if grantee_key in grantees_with_template:
                continue
                
            if grantee_key not in grantee_permissions:
                grantee_permissions[grantee_key] = {
                    'grantee_type': perm['grantee_type'],
                    'grantee_id': perm['grantee_id'],
                    'capabilities': {}
                }
            for cap in perm['capabilities']:
                cap_name = cap['name']
                cap_mode = cap['mode']
                
                # Check if this is a template capability
                if cap_name in workbook_templates:
                    # Expand template
                    for expanded_cap, expanded_mode in workbook_templates[cap_name].items():
                        grantee_permissions[grantee_key]['capabilities'][expanded_cap] = expanded_mode
                else:
                    grantee_permissions[grantee_key]['capabilities'][f"workbook_{cap_name}"] = cap_mode
        
        # Datasource default permissions
        for perm in datasource_perms:
            grantee_key = (perm['grantee_type'], perm['grantee_id'])
            
            # Skip if this grantee already has template applied
            if grantee_key in grantees_with_template:
                continue
                
            if grantee_key not in grantee_permissions:
                grantee_permissions[grantee_key] = {
                    'grantee_type': perm['grantee_type'],
                    'grantee_id': perm['grantee_id'],
                    'capabilities': {}
                }
            for cap in perm['capabilities']:
                cap_name = cap['name']
                cap_mode = cap['mode']
                
                if cap_name in datasource_templates:
                    for expanded_cap, expanded_mode in datasource_templates[cap_name].items():
                        grantee_permissions[grantee_key]['capabilities'][expanded_cap] = expanded_mode
                else:
                    grantee_permissions[grantee_key]['capabilities'][f"datasource_{cap_name}"] = cap_mode
        
        # Flow default permissions
        for perm in flow_perms:
            grantee_key = (perm['grantee_type'], perm['grantee_id'])
            
            # Skip if this grantee already has template applied
            if grantee_key in grantees_with_template:
                continue
                
            if grantee_key not in grantee_permissions:
                grantee_permissions[grantee_key] = {
                    'grantee_type': perm['grantee_type'],
                    'grantee_id': perm['grantee_id'],
                    'capabilities': {}
                }
            for cap in perm['capabilities']:
                cap_name = cap['name']
                cap_mode = cap['mode']
                
                if cap_name in flow_templates:
                    for expanded_cap, expanded_mode in flow_templates[cap_name].items():
                        grantee_permissions[grantee_key]['capabilities'][expanded_cap] = expanded_mode
                else:
                    grantee_permissions[grantee_key]['capabilities'][f"flow_{cap_name}"] = cap_mode
        
        # Virtual connection default permissions
        for perm in virtualconnection_perms:
            grantee_key = (perm['grantee_type'], perm['grantee_id'])
            
            # Track that this grantee has virtualconnection permissions
            grantees_with_virtualconnection.add(grantee_key)
            
            # Skip if this grantee already has template applied
            if grantee_key in grantees_with_template:
                continue
                
            if grantee_key not in grantee_permissions:
                grantee_permissions[grantee_key] = {
                    'grantee_type': perm['grantee_type'],
                    'grantee_id': perm['grantee_id'],
                    'capabilities': {}
                }
            for cap in perm['capabilities']:
                cap_name = cap['name']
                cap_mode = cap['mode']
                
                if cap_name in virtualconnection_templates:
                    for expanded_cap, expanded_mode in virtualconnection_templates[cap_name].items():
                        grantee_permissions[grantee_key]['capabilities'][expanded_cap] = expanded_mode
                else:
                    grantee_permissions[grantee_key]['capabilities'][f"virtualconnection_{cap_name}"] = cap_mode
        
        # Database default permissions
        for perm in database_perms:
            grantee_key = (perm['grantee_type'], perm['grantee_id'])
            
            # Skip if this grantee already has template applied
            if grantee_key in grantees_with_template:
                continue
                
            if grantee_key not in grantee_permissions:
                grantee_permissions[grantee_key] = {
                    'grantee_type': perm['grantee_type'],
                    'grantee_id': perm['grantee_id'],
                    'capabilities': {}
                }
            for cap in perm['capabilities']:
                cap_name = cap['name']
                cap_mode = cap['mode']
                
                if cap_name in database_templates:
                    for expanded_cap, expanded_mode in database_templates[cap_name].items():
                        grantee_permissions[grantee_key]['capabilities'][expanded_cap] = expanded_mode
                else:
                    grantee_permissions[grantee_key]['capabilities'][f"database_{cap_name}"] = cap_mode
        
        # Table default permissions
        for perm in table_perms:
            grantee_key = (perm['grantee_type'], perm['grantee_id'])
            
            # Skip if this grantee already has template applied
            if grantee_key in grantees_with_template:
                continue
                
            if grantee_key not in grantee_permissions:
                grantee_permissions[grantee_key] = {
                    'grantee_type': perm['grantee_type'],
                    'grantee_id': perm['grantee_id'],
                    'capabilities': {}
                }
            for cap in perm['capabilities']:
                cap_name = cap['name']
                cap_mode = cap['mode']
                
                if cap_name in table_templates:
                    for expanded_cap, expanded_mode in table_templates[cap_name].items():
                        grantee_permissions[grantee_key]['capabilities'][expanded_cap] = expanded_mode
                else:
                    grantee_permissions[grantee_key]['capabilities'][f"table_{cap_name}"] = cap_mode
        
        # Convert back to the expected format with capabilities as list
        merged_permissions = []
        for grantee_key, perm_info in grantee_permissions.items():
            # Ensure virtualconnection_Write is present for all grantees with VC permissions (API limitation)
            # Check both: tracked set AND if any VC capability exists
            has_vc_caps = any(cap.startswith('virtualconnection_') for cap in perm_info['capabilities'])
            if grantee_key in grantees_with_virtualconnection or has_vc_caps:
                if 'virtualconnection_Write' not in perm_info['capabilities']:
                    perm_info['capabilities']['virtualconnection_Write'] = 'No API Endpoint'
            
            perm_entry = {
                'grantee_type': perm_info['grantee_type'],
                'grantee_id': perm_info['grantee_id'],
                'asset_permissions': content_permissions,
                'capabilities': []
            }
            # Convert capabilities dict back to list format
            for cap_name, cap_mode in perm_info['capabilities'].items():
                perm_entry['capabilities'].append({
                    'name': cap_name,
                    'mode': cap_mode
                })
            merged_permissions.append(perm_entry)
        
        return merged_permissions
    
    def _get_default_permissions(self, project_id: str, content_type: str) -> List[Dict]:
        """Get project default permissions for a content type"""
        url = f"{self.base_url}/sites/{self.site_id}/projects/{project_id}/default-permissions/{content_type}"
        try:
            response = requests.get(url, headers=self._get_headers())
            response.raise_for_status()
            return self._parse_permissions(response.content)
        except:
            return []
    
    def export_permissions_to_csv(self, permissions: List[Dict], filename: str):
        """Export Permission Rules to CSV with GUI-friendly headers"""
        if not permissions:
            return
        
        # API to GUI name mapping (common)
        api_to_gui_name_common = {
            'Read': 'View',
            'ExportImage': 'Download Image/PDF',
            'ExportData': 'Download Summary Data',
            'ShareView': 'Share Customised',
            'ViewUnderlyingData': 'Download Full Data',
            'WebAuthoring': 'Web Edit',
            'Write': 'Overwrite',
            'ChangeHierarchy': 'Move',
            'ChangePermissions': 'Set Permissions',
            'CreateRefreshMetrics': 'Extract Refresh',
            'ExtractRefresh': 'Extract Refresh',
            'SaveAs': 'Save As',
            'InheritedProjectLeader': 'Publish',
            'VizqlDataApiAccess': 'API Access',
            'PulseMetricDefine': 'Create Metric Definitions',
            'Connect': 'Connect',
            'Filter': 'Filter',
            'ViewComments': 'View Comments',
            'AddComment': 'Add Comments',
            'RunExplainData': 'Run Explain Data',
            'Delete': 'Delete',
            'ProjectLeader': 'Project Leader',
            'Execute': 'Run Flow',
            'WebAuthoringForFlows': 'Web Edit',
            'CreateVirtualConnection': 'Publish',
            'CreateDatabase': 'Publish',
            'CreateTable': 'Publish',
        }
        
        # Content type specific mappings
        api_to_gui_name_by_type = {
            'workbook': {'ExportXml': 'Download/Save a Copy'},
            'datasource': {'ExportXml': 'Download Data Source'},
            'view': {'ExportXml': 'Download/Save a Copy'},
            'flow': {'ExportXml': 'Download Flow'},
            'project': {'Write': 'Publish'},
        }
        
        # Capability ordering by content type
        capability_order = {
            'workbook': [
                'Read', 'Filter', 'ViewComments', 'AddComment', 'ExportImage', 'ExportData',
                'ShareView', 'ViewUnderlyingData', 'WebAuthoring', 'RunExplainData', 'ExportXml',
                'Write', 'ChangeHierarchy', 'Delete', 'ChangePermissions', 'ExtractRefresh',
            ],
            'datasource': [
                'Read', 'Connect', 'ExportXml', 'Write', 'SaveAs', 'VizqlDataApiAccess',
                'PulseMetricDefine', 'ChangeHierarchy', 'Delete', 'ChangePermissions', 'ExtractRefresh',
            ],
            'project': [
                'Read', 'Write',
                'workbook_Read', 'workbook_Filter', 'workbook_ViewComments', 'workbook_AddComment',
                'workbook_ExportImage', 'workbook_ExportData', 'workbook_ShareView',
                'workbook_ViewUnderlyingData', 'workbook_WebAuthoring', 'workbook_RunExplainData',
                'workbook_ExportXml', 'workbook_Write', 'workbook_ChangeHierarchy', 'workbook_Delete',
                'workbook_ChangePermissions', 'workbook_ExtractRefresh',
                'datasource_Read', 'datasource_Connect', 'datasource_ExportXml', 'datasource_Write',
                'datasource_SaveAs', 'datasource_VizqlDataApiAccess', 'datasource_PulseMetricDefine',
                'datasource_ChangeHierarchy', 'datasource_Delete', 'datasource_ChangePermissions',
                'datasource_ExtractRefresh',
                'flow_Read', 'flow_ExportXml', 'flow_Execute', 'flow_Write', 'flow_WebAuthoringForFlows',
                'flow_ChangeHierarchy', 'flow_Delete', 'flow_ChangePermissions',
                'virtualconnection_Read', 'virtualconnection_Connect', 'virtualconnection_Write',
                'virtualconnection_ChangeHierarchy', 'virtualconnection_Delete', 'virtualconnection_ChangePermissions',
                'database_Read', 'database_Write', 'database_ChangeHierarchy', 'database_ChangePermissions',
                'table_Read', 'table_Write', 'table_ChangeHierarchy', 'table_ChangePermissions',
            ],
            'view': [
                'Read', 'Filter', 'ViewComments', 'AddComment', 'ExportImage', 'ExportData',
                'ShareView', 'ViewUnderlyingData', 'WebAuthoring', 'Delete', 'ChangePermissions'
            ],
            'flow': [
                'Read', 'ExportXml', 'Execute', 'Write', 'WebAuthoringForFlows',
                'ChangeHierarchy', 'Delete', 'ChangePermissions'
            ]
        }
        
        # Collect all fields from actual data
        all_fields = set()
        all_capabilities_in_data = set()
        for perm in permissions:
            all_fields.update(perm.keys())
            for key in perm.keys():
                if key not in ['content_type', 'content_id', 'content_name', 'workbook_name', 'project_name',
                              'parent_project_id', 'asset_permissions', 'grantee_type', 'grantee_id', 'grantee_name']:
                    all_capabilities_in_data.add(key)
        
        # Remove duplicates
        if 'ExtractRefresh' in all_capabilities_in_data and 'CreateRefreshMetrics' in all_capabilities_in_data:
            all_capabilities_in_data.discard('CreateRefreshMetrics')
        if 'VizqlDataApiAccess' in all_capabilities_in_data and 'InheritedProjectLeader' in all_capabilities_in_data:
            all_capabilities_in_data.discard('InheritedProjectLeader')
        if 'workbook_ExtractRefresh' in all_capabilities_in_data and 'workbook_CreateRefreshMetrics' in all_capabilities_in_data:
            all_capabilities_in_data.discard('workbook_CreateRefreshMetrics')
        if 'datasource_ExtractRefresh' in all_capabilities_in_data and 'datasource_CreateRefreshMetrics' in all_capabilities_in_data:
            all_capabilities_in_data.discard('datasource_CreateRefreshMetrics')
        if 'datasource_VizqlDataApiAccess' in all_capabilities_in_data and 'datasource_InheritedProjectLeader' in all_capabilities_in_data:
            all_capabilities_in_data.discard('datasource_InheritedProjectLeader')
        
        # Define fixed fields
        fixed_fields = ['content_type', 'content_id', 'content_name', 'workbook_name', 'project_name',
                       'parent_project_id', 'asset_permissions', 'grantee_type', 'grantee_id', 'grantee_name']
        
        # Determine content type
        content_types_in_data = set(perm.get('content_type', 'workbook') for perm in permissions)
        
        if len(content_types_in_data) == 1:
            content_type = list(content_types_in_data)[0]
        else:
            if 'workbook' in content_types_in_data:
                content_type = 'workbook'
            else:
                content_type = list(content_types_in_data)[0]
        
        # Get ALL capabilities for this content type (not just ones in data)
        ordered_capabilities = capability_order.get(content_type, [])
        
        # Use ALL capabilities from the order, even if not in data
        all_capabilities = set(ordered_capabilities)
        # Also include any capabilities that are in data but not in order
        all_capabilities.update(all_capabilities_in_data)
        
        # Sort capabilities
        sorted_capabilities = []
        for cap in ordered_capabilities:
            sorted_capabilities.append(cap)
        
        # Add remaining capabilities that weren't in the predefined order
        remaining_caps = sorted(all_capabilities - set(ordered_capabilities))
        sorted_capabilities.extend(remaining_caps)
        
        # Create display headers with GUI names (removed "Name:" prefix)
        display_headers = []
        for cap in sorted_capabilities:
            gui_name = None
            
            if content_type == 'project':
                if cap.startswith('workbook_'):
                    base_cap = cap.replace('workbook_', '')
                    if base_cap in api_to_gui_name_by_type.get('workbook', {}):
                        gui_name = api_to_gui_name_by_type['workbook'][base_cap]
                    elif base_cap in api_to_gui_name_common:
                        gui_name = api_to_gui_name_common[base_cap]
                    if gui_name:
                        display_headers.append(f"Workbooks Tab - {gui_name}(API Endpoints:{base_cap})")
                    else:
                        display_headers.append(f"Workbooks Tab - {base_cap}")
                    continue
                elif cap.startswith('datasource_'):
                    base_cap = cap.replace('datasource_', '')
                    if base_cap in api_to_gui_name_by_type.get('datasource', {}):
                        gui_name = api_to_gui_name_by_type['datasource'][base_cap]
                    elif base_cap in api_to_gui_name_common:
                        gui_name = api_to_gui_name_common[base_cap]
                    if gui_name:
                        display_headers.append(f"Data Sources Tab - {gui_name}(API Endpoints:{base_cap})")
                    else:
                        display_headers.append(f"Data Sources Tab - {base_cap}")
                    continue
                elif cap.startswith('flow_'):
                    base_cap = cap.replace('flow_', '')
                    if base_cap in api_to_gui_name_by_type.get('flow', {}):
                        gui_name = api_to_gui_name_by_type['flow'][base_cap]
                    elif base_cap in api_to_gui_name_common:
                        gui_name = api_to_gui_name_common[base_cap]
                    if gui_name:
                        display_headers.append(f"Flows Tab - {gui_name}(API Endpoints:{base_cap})")
                    else:
                        display_headers.append(f"Flows Tab - {base_cap}")
                    continue
                elif cap.startswith('virtualconnection_'):
                    base_cap = cap.replace('virtualconnection_', '')
                    if base_cap == 'Write':
                        display_headers.append(f"Virtual Connections Tab - Overwrite(API Endpoints:No API Endpoints)")
                        continue
                    if base_cap in api_to_gui_name_common:
                        gui_name = api_to_gui_name_common[base_cap]
                    if gui_name:
                        display_headers.append(f"Virtual Connections Tab - {gui_name}(API Endpoints:{base_cap})")
                    else:
                        display_headers.append(f"Virtual Connections Tab - {base_cap}")
                    continue
                elif cap.startswith('database_'):
                    base_cap = cap.replace('database_', '')
                    if base_cap in api_to_gui_name_common:
                        gui_name = api_to_gui_name_common[base_cap]
                    if gui_name:
                        display_headers.append(f"Databases Tab - {gui_name}(API Endpoints:{base_cap})")
                    else:
                        display_headers.append(f"Databases Tab - {base_cap}")
                    continue
                elif cap.startswith('table_'):
                    base_cap = cap.replace('table_', '')
                    if base_cap in api_to_gui_name_common:
                        gui_name = api_to_gui_name_common[base_cap]
                    if gui_name:
                        display_headers.append(f"Tables Tab - {gui_name}(API Endpoints:{base_cap})")
                    else:
                        display_headers.append(f"Tables Tab - {base_cap}")
                    continue
                else:
                    if cap in api_to_gui_name_by_type.get('project', {}):
                        gui_name = api_to_gui_name_by_type['project'][cap]
                        display_headers.append(f"Projects Tab - {gui_name}(API Endpoints:{cap})")
                    elif cap in api_to_gui_name_common:
                        gui_name = api_to_gui_name_common[cap]
                        display_headers.append(f"Projects Tab - {gui_name}(API Endpoints:{cap})")
                    else:
                        display_headers.append(f"Projects Tab - {cap}")
                    continue
            
            if content_type in api_to_gui_name_by_type and cap in api_to_gui_name_by_type[content_type]:
                gui_name = api_to_gui_name_by_type[content_type][cap]
            elif cap in api_to_gui_name_common:
                gui_name = api_to_gui_name_common[cap]
            
            if gui_name:
                display_headers.append(f"{gui_name}(API Endpoints:{cap})")
            else:
                display_headers.append(cap)
        
        fieldnames = [f for f in fixed_fields if f in all_fields]
        
        # Write CSV
        with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(fieldnames + display_headers)
            
            for perm in permissions:
                row = [perm.get(field, '') for field in fieldnames]
                for cap in sorted_capabilities:
                    value = perm.get(cap, '')
                    # Special handling for virtualconnection_Write (API limitation)
                    if cap == 'virtualconnection_Write' and value == '':
                        value = 'No API Endpoint'
                    # Replace empty values with "No setting"
                    elif value == '':
                        value = 'No setting'
                    row.append(value)
                writer.writerow(row)
    
    def sign_out(self):
        """Sign out"""
        if not self.auth_token:
            return
        
        url = f"{self.base_url}/auth/signout"
        try:
            requests.post(url, headers=self._get_headers())
        except:
            pass
        
        self.auth_token = None
        self.site_id = None
        self.user_id = None
