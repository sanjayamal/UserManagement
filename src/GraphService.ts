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
  await client.api(`/users/${id}`).delete()
              .then((res:any) => res);
  //return res;
}

export async function updateUser(accessToken:string,user:any){
   const {id}=user;
   console.log(id)
   const client = getAuthenticatedClient(accessToken);
   await client.api(`/users/${id}`).update(user)
                .then((res:any) => res);
  //return res;
}