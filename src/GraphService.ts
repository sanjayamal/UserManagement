
var graph = require('@microsoft/microsoft-graph-client');

function getAuthenticatedClient(accessToken: string) {
  // Initialize Graph client
  const client = graph.Client.init({
    // Use the provided access token to authenticate
    // requests
    authProvider: (done: any) => {
      done(null, accessToken);
    }
  });

  return client;
}

export async function getEvents(accessToken: string) {
    const client = getAuthenticatedClient(accessToken);
  
    const events = await client
      .api('/me/events')
      .select('subject,organizer,start,end')
      .orderby('createdDateTime DESC')
      .get();
  
    return events;
  }
  
export async function getUserDetails(accessToken: string) {
  const client = getAuthenticatedClient(accessToken);

  const user = await client.api('/me').get();
  return user;
}

export async function createUser(accessToken:string,user:any){
  const client = getAuthenticatedClient(accessToken);
  await client.api('/users').post(user)
  .then((res:any)=>res);
 // return res;
}


export async function getUser(accessToken:string){
  const client = getAuthenticatedClient(accessToken);
  let res = await client.api('/users').get();
  return res;
}

export async function deleteUser(accessToken:string,id:any){
  const client = getAuthenticatedClient(accessToken);
  let res=await client.api(`/users/${id}`).delete()
  return res;
}

export async function updateUser(accessToken:string,user:any){
   const {id}=user;
   console.log(id)
   const client = getAuthenticatedClient(accessToken);
   let res= await client.api(`/users/${id}`).update(user)
  return res;
}

export async function addUserGroup(accessToken:string,groupId:string,user:any){
  
  const client = getAuthenticatedClient(accessToken);
  let res = await client.api(`/groups/${groupId}`)
  .update(user);
  // let res = await client.api(`/groups/${groupId}/members/28c4665d-a57b-4dfb-bb1f-6a0e6e52f1a5/$ref`)
	// .delete();
  return res;
}

export async function getGroup(accessToken:string){
  const client = getAuthenticatedClient(accessToken);
  let res = await client.api('/groups')
  .get();
  return res;
}

export async function getMemberGroups(accessToken:string,id:string,reqBody:any){
  const client = getAuthenticatedClient(accessToken);
  let res = await client.api(`/users/${id}/getMemberGroups`)
  .post(reqBody);
  return res;
}

export async function deleteUserGroup(accessToken:string,groupId:string,userId:any){
  const client = getAuthenticatedClient(accessToken);
  let res = await client.api(`/groups/${groupId}/members/${userId}/$ref`)
	.delete();
  return res;
}