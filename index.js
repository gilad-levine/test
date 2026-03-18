const hubspot = require('@hubspot/api-client');
exports.main = async (event, callback) => {
  const hubspotClient = new hubspot.Client({
    accessToken: process.env.TestSecret
  });
  const SUBMISSION_OBJECT_TYPE = '2-52975400';
  try {
    const contactId = String(event.inputFields['hs_object_id']);
    console.log('Contact ID:', contactId);
    const submissionAssocResp = await hubspotClient.apiRequest({
      method: 'GET',
      path: `/crm/v4/objects/contacts/${contactId}/associations/${SUBMISSION_OBJECT_TYPE}`
    });
    const submissionAssocData = await submissionAssocResp.json();
    if (!submissionAssocData.results || submissionAssocData.results.length === 0) {
      throw new Error('No submissions associated with contact ' + contactId);
    }
    const submissions = await Promise.all(
      submissionAssocData.results.map(async (r) => {
        const obj = await hubspotClient.crm.objects.basicApi.getById(
          SUBMISSION_OBJECT_TYPE,
          String(r.toObjectId),
          ['hs_createdate', 'submission_type']
        );
        return {
          id: r.toObjectId,
          createdate: obj.properties.hs_createdate,
          submission_type: obj.properties.submission_type
        };
      })
    );
    const registrationSubmissions = submissions.filter(
      (s) => s.submission_type === 'Registration'
    );
    if (registrationSubmissions.length === 0) {
      throw new Error('No Registration submissions associated with contact ' + contactId);
    }
    const sortedByDate = registrationSubmissions.sort(
      (a, b) => new Date(b.createdate) - new Date(a.createdate)
    );
    const submissionId = String(sortedByDate[0].id);
    console.log('Most recent submission ID:', submissionId);
    const submissionResp = await hubspotClient.apiRequest({
      method: 'GET',
      path: `/crm/v3/objects/${SUBMISSION_OBJECT_TYPE}/${submissionId}?properties=event,submission_name`
    });
    const submissionData = await submissionResp.json();
    const eventValue = submissionData.properties?.event || '';
    const submissionName = submissionData.properties?.submission_name || '';
    console.log('Event:', eventValue);
    console.log('Submission name:', submissionName);
    callback({
      outputFields: {
        event: eventValue,
        submission_id: submissionId,
        submission_name: submissionName
      }
    });
  } catch (error) {
    console.error('Error:', error.message);
    callback({
      outputFields: {
        event: '',
        submission_id: '',
        submission_name: ''
      }
    });
  }
};
