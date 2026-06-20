export const PRTSService = {
  async getRequest(siteUrl: string, query: string) {
    const response = await fetch(siteUrl + query, {
      method: "GET",
      headers: {
        Accept: "application/json;odata=verbose",
      },
    });

    if (!response.ok) {
      throw new Error(`Error fetching data: ${response.statusText}`);
    }

    const data = await response.json();
    return data.d?.results || data.d;
  },

  async GetSPListItems(siteUrl: string, site: string, listName: string) {
    const query = `/${site}/_api/web/lists/GetByTitle('${listName}')/items?$select=ID,Title`;
    return this.getRequest(siteUrl, query);
  },

  async GetSPUserDetailsListItems(siteUrl: string, userEmail: string) {
    const query = `/_api/web/lists/GetByTitle('EmployeeProfile')/items?$select=ID,EmpName/Title,EmpName/EMail,EmpDepartment/Title&$expand=EmpName,EmpDepartment&$filter=EmpName/EMail eq '${userEmail}'`;
    return this.getRequest(siteUrl, query);
  },

  async AttachmentsINListItems(siteUrl: string, id: number) {
    const query = `/_api/web/lists/GetByTitle('PRTSList')/items('${id}')/AttachmentFiles`;
    return this.getRequest(siteUrl, query);
  },

  async GetDocumentTemplates(siteUrl: string) {
    const query = `/_api/web/lists/GetByTitle('Template_DocumentList')/items?$select=Title,filepath`;
    return this.getRequest(siteUrl, query);
  },

  async GetSPListDiamondWF_UserName(siteUrl: string, area: string) {
    const query = `/_api/web/lists/GetByTitle('PRTS_7D_Matrix')/items?$select=Title,UserDiamond1/Title,UserDiamond2/Title,UserDiamond3/Title,UserDiamond4/Title,UserDiamond5/Title,UserDiamond6/Title,UserDiamond7/Title&$expand=UserDiamond1,UserDiamond2,UserDiamond3,UserDiamond4,UserDiamond5,UserDiamond6,UserDiamond7&$filter= Title eq '${area}'`;
    return this.getRequest(siteUrl, query);
  },

  async GetSPListDelegation(siteUrl: string) {
    const query = `/_api/web/lists/GetByTitle('Delegation')/items?$select=DelegateFrom/Title,DelegateTo/Title,Status,AppList&$expand=DelegateFrom,DelegateTo&$filter=Status eq 'ON'`;
    return this.getRequest(siteUrl, query);
  },
};
