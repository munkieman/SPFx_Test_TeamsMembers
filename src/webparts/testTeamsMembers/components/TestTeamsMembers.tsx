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
  const [error, setError] = useState<string | null>(null);
  const teamName = "Travel and Expenses";
  const channelName = "ExpensesChat";

  //componentDidMount
  useEffect(() => {
    console.log("componentDidMount called.");

    const fetchChannelMembers = async () => {
      try {
        setLoading(true);

        // Get Microsoft Graph API client
        const client = await context.aadHttpClientFactory.getClient("https://graph.microsoft.com");

        // Get team ID
        const teamsResponse: HttpClientResponse = await client.get(
          `https://graph.microsoft.com/v1.0/me/joinedTeams`,
          AadHttpClient.configurations.v1
        );

        if (!teamsResponse.ok) throw new Error("Failed to fetch teams");

        const teamsData = await teamsResponse.json();
        const team = teamsData.value.find((t: any) => t.displayName === teamName);
        if (!team) throw new Error(`Team "${teamName}" not found`);

        // Get channel ID
        const channelsResponse: HttpClientResponse = await client.get(
          `https://graph.microsoft.com/v1.0/teams/${team.id}/channels`,
          AadHttpClient.configurations.v1
        );

        if (!channelsResponse.ok) throw new Error("Failed to fetch channels");

        const channelsData = await channelsResponse.json();
        const channel = channelsData.value.find((c: any) => c.displayName === channelName);
        if (!channel) throw new Error(`Channel "${channelName}" not found`);

        // Get channel members
        const membersResponse: HttpClientResponse = await client.get(
          `https://graph.microsoft.com/v1.0/teams/${team.id}/channels/${channel.id}/members`,
          AadHttpClient.configurations.v1
        );

        if (!membersResponse.ok) throw new Error("Failed to fetch channel members");

        const membersData = await membersResponse.json();
        let membersList: IChannelMember[] = membersData.value.map((m: any) => ({
          id: m.userId,
          displayName: m.displayName,
          roles: m.roles.map((r: string) => r.toLowerCase()), // Normalize roles
        }));
        // 🔥 Remove Owners from the list
        membersList = membersList.filter((member) => !member.roles.includes("owner"));


        // Fetch presence for each remaining member
        const membersWithPresence = await Promise.all(
          membersList.map(async (member) => {
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
        //setMembers(membersData.value);
      } catch (err: any) {
        setError(err.message);
      } finally {
        setLoading(false);
      }    
    };

    fetchChannelMembers(); 
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
            <strong> {member.presence} </strong>
            </li>
          ))}
        </ul>
      </div>
    </section>
  );
}
export default TestTeamsMembers;