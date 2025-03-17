import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './TestTeamsMembers.module.scss';
import type { ITestTeamsMembersProps } from './ITestTeamsMembersProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { AadHttpClient, HttpClientResponse } from "@microsoft/sp-http";

interface IChannelMember {
  id: string;
  displayName: string;
  roles: string[];
  presence?: string; // Online, Busy, etc.
}

interface DialogProps {
  type: "error" | "success" | "warning" | "info";
  message: string | null;
  onClose: () => void;
}

//https://teams.microsoft.com/l/channel/19%3A4MJijyhTk1dVGDiXO5LOCHxxMqv2Iz-wD6Wtco1W7j81%40thread.tacv2/General?groupId=696dfe67-e76f-4bf8-8ab6-8abfcb16552e&tenantId=1a25c064-c00a-402f-8f6c-ce0e12a6293d

const TestTeamsMembers: React.FunctionComponent<ITestTeamsMembersProps> = (props:ITestTeamsMembersProps) => {

  const {
    description,
    isDarkTheme,
    environmentMessage,
    hasTeamsContext,
    userDisplayName,
    context 
  } = props;

  const [members, setMembers] = useState<IChannelMember[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [message, setMessage] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [dialogMessage, setDialogMessage] = useState<string | null>(null);
  const [isDialogOpen, setIsDialogOpen] = useState<boolean>(false);
  const [dialogType, setDialogType] = useState<"error" | "success" | "warning" | "info">("info"); // Default type
  const teamName = "TestChat";
  const channelName = "General";

  const Dialog: React.FC<DialogProps> = ({ type, message, onClose }) => {
    if (!message) return null;
  
    return (
      <div className={styles.dialogOverlay}>
        <div className={`${styles.dialogBox} ${styles[type]}`}>
          <p>{message}</p>
          <button onClick={onClose}>Close</button>
        </div>
      </div>
    );
  };

  const showDialog = (type: "error" | "success" | "warning" | "info", message: string) => {
    setDialogType(type);
    setDialogMessage(message);
    setIsDialogOpen(true);
  };

  //componentDidMount
  useEffect(() => {
    console.log("componentDidMount called.");

    const fetchTeamMembers = async () => {
      try {
        setLoading(true);
    
        // Get Microsoft Graph API client
        const client = await context.aadHttpClientFactory.getClient("https://graph.microsoft.com");
    
        // Get team ID
        const teamsResponse: HttpClientResponse = await client.get(
          `https://graph.microsoft.com/v1.0/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')`,
          AadHttpClient.configurations.v1
        );
    
        if (!teamsResponse.ok) throw new Error("Failed to fetch teams");
    
        const teamsData = await teamsResponse.json();
        const team = teamsData.value.find((t: any) => t.displayName === teamName);
        if (!team) throw new Error(`Team "${teamName}" not found`);
    
        console.log("Found Team:", team);
    
        // 🔹 Step 1: Fetch All Team Tags
        const tagsResponse: HttpClientResponse = await client.get(
          `https://graph.microsoft.com/v1.0/teams/${team.id}/tags`,
          AadHttpClient.configurations.v1
        );

        if (!tagsResponse.ok) throw new Error("Failed to fetch team tags");

        const tagsData = await tagsResponse.json();
        const expensesTag = tagsData.value.find((tag: any) => tag.displayName.toLowerCase() === "expenses");

        if (!expensesTag) throw new Error('Tag "expenses" not found in the team');

        console.log("Found Tag:", expensesTag);

        // 🔹 Step 2: Get Members Assigned to "expenses" Tag
        const tagMembersResponse: HttpClientResponse = await client.get(
          `https://graph.microsoft.com/v1.0/teams/${team.id}/tags/${expensesTag.id}/members`,
          AadHttpClient.configurations.v1
        );

        if (!tagMembersResponse.ok) throw new Error('Failed to fetch members with "expenses" tag');

        const tagMembersData = await tagMembersResponse.json();
        const tagMemberIds = tagMembersData.value.map((m: any) => m.userId);
    
        console.log("Tag Members:", tagMemberIds);

        // 🔹 Step 3: Get All Team Members
        // Get standard channel ID (assuming General channel)        
        const channelsResponse: HttpClientResponse = await client.get(
          `https://graph.microsoft.com/v1.0/teams/${team.id}/channels`,
          AadHttpClient.configurations.v1
        );
    
        if (!channelsResponse.ok) throw new Error("Failed to fetch channels");
    
        const channelsData = await channelsResponse.json();
        const channel = channelsData.value.find((c: any) => c.displayName === channelName); 
        if (!channel) throw new Error(`Channel "General" not found in team "${teamName}"`);
    
        console.log("Found Channel:", channel);
    
        // Get channel members
        const membersResponse: HttpClientResponse = await client.get(
          `https://graph.microsoft.com/v1.0/teams/${team.id}/members`,
          AadHttpClient.configurations.v1
        );
    
        if (!membersResponse.ok) throw new Error("Failed to fetch team members");
    
        const membersData = await membersResponse.json();
    
        // 🔹 Step 4: Process Members
        let membersList = membersData.value
          .map((m: any) => ({
            id: m.userId,
            displayName: m.displayName,
            roles: Array.isArray(m.roles) ? m.roles.map((r: string) => r.toLowerCase()) : [],
          }))
          .filter((member: { id: any; }) => tagMemberIds.includes(member.id)) // Filter by tag members
          .filter((member: { roles: string | string[]; }) => !member.roles.includes("owner")); // Remove Owners

        // Process members
        /*
        let membersList = membersData.value.map((m: any) => ({
          id: m.userId,
          displayName: m.displayName,
          roles: m.roles.map((r: string) => r.toLowerCase()), // Normalize roles
        }));
        */
    
        // 🔥 Remove Owners from the list
        membersList = membersList.filter((member: { roles: string | string[]; }) => !member.roles.includes("owner"));
    
        // 🔹 Step 5: Fetch Presence for Each Member
        const membersWithPresence = await Promise.all(
          membersList.map(async (member: { id: any; }) => {
            try {
              const presenceResponse: HttpClientResponse = await client.get(
                `https://graph.microsoft.com/v1.0/users/${member.id}/presence`,
                AadHttpClient.configurations.v1
              );
              if (!presenceResponse.ok) throw new Error("Failed to fetch presence");
    
              const presenceData = await presenceResponse.json();
              return { ...member, presence: presenceData.availability };
            } catch {
              return { ...member, presence: "Unknown" };
            }
          })
        );
    
        setMembers(membersWithPresence);
      } catch (err: any) {
        console.error("Error fetching team members:", err.message);
        setError(err.message);
      } finally {
        setLoading(false);
      }
    };

    const checkMember = async () => { 
      try {
        //const accessToken = await getAccessToken();
        const client = await context.aadHttpClientFactory.getClient("https://graph.microsoft.com");
        const userResponse = await client.get(
          `https://graph.microsoft.com/v1.0/users/${props.context.pageContext.user.email}`,
          AadHttpClient.configurations.v1
        );
        const userData = await userResponse.json();
        const userId = userData.id;      
        console.log("userID",userId,userData);

        // Fetch Team ID
        const teamsResponse: HttpClientResponse = await client.get(
          `https://graph.microsoft.com/v1.0/me/joinedTeams`,
          AadHttpClient.configurations.v1
        );
        if (!teamsResponse.ok) throw new Error("Failed to fetch teams");
    
        const teamsData = await teamsResponse.json();
        let team = teamsData.value.find((t: any) => t.displayName === teamName);
        //if (!team) throw new Error(`Team "${teamName}" not found`);
        
        if (!team) {
          console.log(`User is not in the team "${teamName}". Fetching team ID manually...`);
    
          // Get all teams the user has access to
          const allTeamsResponse = await client.get(
            `https://graph.microsoft.com/v1.0/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')`,
            AadHttpClient.configurations.v1
          );
          if (!allTeamsResponse.ok) throw new Error("Failed to fetch available teams");
    
          const allTeamsData = await allTeamsResponse.json();
          team = allTeamsData.value.find((t: any) => t.displayName === teamName);
    
          if (!team) throw new Error(`Team "${teamName}" not found.`);
        }
    
        console.log("Found Team:", team);
        console.log("checkmember Team:", team);        

        // Check if user is already a member of the team
        const membersResponse = await client.get(
          `https://graph.microsoft.com/v1.0/teams/${team.id}/members`,
          AadHttpClient.configurations.v1
        );
        const membersData = await membersResponse.json();
        const userIsMember = membersData.value.some((m: any) => m.id === userId)
        //const userIsMember = members.some(member => member.id === userId);    
        
        if (!userIsMember) {
          showDialog("info","Adding you to the chat channel. Please wait...");    
          console.log("User is not a member, adding to the chat channel...");
    
    
          // Add user to the team
          const addUserResponse: HttpClientResponse = await client.post(
            `https://graph.microsoft.com/v1.0/teams/${team.id}/members`, //channels/${channel.id}/members`,
             AadHttpClient.configurations.v1,
            {
              headers: { 
                "Content-Type": "application/json"
                //Authorization : `Bearer ${accessToken}`,
              },
              body: JSON.stringify({
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": ["member"],
                "user@odata.bind": `https://graph.microsoft.com/v1.0/users/${userId}`,
                "isHistoryIncluded": false, 
                "visibleHistoryStartDateTime": null
              })
            }
          );
    
          if (!addUserResponse.ok) {
            const errorText = await addUserResponse.text();
            showDialog("error","Failed to add user to the chat: " + errorText);
            throw new Error(`Failed to add user to the chat: ${errorText}`);
          }

          fetchTeamMembers();       
          showDialog("success","You have successfully joined the chat!");

          //setIsChatDisabled(true);
          console.log("User successfully added to the chat");
        }else{
          showDialog("success","You are already a member of this chat channel.");        
          console.log("User is already a member of the chat channel");          
        }
        
        //setChatContent(
        //  `<iframe class="${styles.chatFrame}" src=""https://teams.microsoft.com/embed-client/chats/list?layout=singlePane"`
           //https://teams.microsoft.com/l/channel/19%3AhC7tyJQiEwWgSjdfY12Kog0xog_43X9rEKdeLxxPP681%40thread.tacv2/General?groupId=ce155c65-5e9b-43a3-87c1-dd5ccc2d2fd3&tenantId=60b37d9e-2c27-417c-8f55-d82b676764bf"></iframe>`
        //);
    
      } catch (error: any) {
        console.error("Error in checkMember:", error.message);
        showDialog("error","Error in adding user - checkMember() : " + error.message + ".  Please report this issue to the Service Desk.");
      }
    };  

    const removeMember = async () => {
      alert('Removing member from chat');
      /*
      try {
        const client = await context.aadHttpClientFactory.getClient("https://graph.microsoft.com");
        const userEmail = props.context.pageContext.user.email;
        
        // Fetch Team ID
        const teamsResponse: HttpClientResponse = await client.get(
          `https://graph.microsoft.com/v1.0/me/joinedTeams`,
          AadHttpClient.configurations.v1
        );
    
        if (!teamsResponse.ok) throw new Error("Failed to fetch teams");
    
        const teamsData = await teamsResponse.json();
        const team = teamsData.value.find((t: any) => t.displayName === teamName);
        if (!team) throw new Error(`Team "${teamName}" not found`);
    
        // Fetch Standard Channel ID (not shared channel)
        const channelsResponse: HttpClientResponse = await client.get(
          `https://graph.microsoft.com/v1.0/teams/${team.id}/channels`,
          AadHttpClient.configurations.v1
        );
        if (!channelsResponse.ok) throw new Error("Failed to fetch channels");
    
        const channelsData = await channelsResponse.json();
        const channel = channelsData.value.find((c: any) => c.displayName === channelName);
        if (!channel) throw new Error(`Channel "${channelName}" not found`);
    
        // 🔥 Fetch members from the standard channel
        const membersResponse: HttpClientResponse = await client.get(
          `https://graph.microsoft.com/v1.0/teams/${team.id}/channels/${channel.id}/members`,
          AadHttpClient.configurations.v1
        );
    
        if (!membersResponse.ok) throw new Error("Failed to fetch channel members");
    
        const membersData = await membersResponse.json();
        const userMember = membersData.value.find((m: any) => m.email === userEmail);
    
        if (!userMember) {
          showDialog("warning", "You are not a member of this channel.");
          return;
        }
    
        // 🔥 Remove the user from the standard channel
        const removeResponse: HttpClientResponse = await client.delete(
          `https://graph.microsoft.com/v1.0/teams/${team.id}/channels/${channel.id}/members/${userMember.id}`,
          AadHttpClient.configurations.v1
        );
    
        if (!removeResponse.ok) {
          throw new Error("Failed to remove user from the chat channel");
        }
    
        showDialog("success", "You have successfully left the chat.");
        console.log("User successfully removed from the standard channel");
      } catch (error: any) {
        console.error("Error removing user:", error.message);
        showDialog("error", "Failed to leave chat. Please try again.");
      }
      */      
    }
    
    checkMember();

    return;   
  }, []);

  //componentDidUpdate
  //useEffect(() => {
  //  console.log("componentDidUpdate called.");
  //}, [count]);

  //componentWillUnmount
  //useEffect(() => {
  //  return () => {
  //    console.log("componentWillUnmount called.");
  //  };
  //}, [count]);

  return (
    <section className={`${styles.testTeamsMembers} ${hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.welcome}>
        <div>
          {isDialogOpen && <Dialog type={dialogType} message={dialogMessage} onClose={() => setIsDialogOpen(false)} />}
        </div>

        <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
        <h2>Well done, {escape(userDisplayName)}!</h2>
        <div>{environmentMessage}</div>
        <div>Web part property value: <strong>{escape(description)}</strong></div>
      </div>
      <div>
        <h3>Team Members & Status (Excluding Owners)</h3>
        {loading && <p>Loading members...</p>}
        {error && <p style={{ color: "red" }}>{error}</p>}
        <ul>
          {members.map((member) => (
            <li key={member.id}>
            {member.displayName} ({member.roles.length > 0 ? member.roles.join(", ") : "Member"}) - 
            <strong style={{ color: member.presence === "Available" ? "green" : "black" }}> {member.presence} </strong>
            </li>
          ))}
        </ul>
        <button className={styles.chatButton} id="chatButton" disabled={!message.trim()}>Join Chat</button>
      </div>
    </section>
  );
}
export default TestTeamsMembers;