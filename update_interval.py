import sys

filepath = 'Code.gs'
with open(filepath, 'r') as f:
    content = f.read()

search_str = """
    // 1. Build a whitelist of "Email" Channel Agents from the Masterlist (With Cache)
    const cache = CacheService.getScriptCache();
    let emailAgentsList = cache.get('emailAgentsList');
    let emailAgents = new Set();

    if (emailAgentsList) {
      emailAgents = new Set(JSON.parse(emailAgentsList));
    } else if (masterSheet && masterSheet.getLastRow() > 1) {
      const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 24).getValues();
      for (let i = 0; i < masterData.length; i++) {
        const ldap = masterData[i][0] ? masterData[i][0].toString().toLowerCase() : '';
        const channel = masterData[i][23] ? masterData[i][23].toString().toLowerCase() : '';

        if (channel === 'email') {
          emailAgents.add(ldap);
        }
      }
      cache.put('emailAgentsList', JSON.stringify(Array.from(emailAgents)), 14400);
    }
"""

replace_str = """
    // 1. Build a whitelist of "Email" Channel Agents & cache Demographics from the Masterlist
    const cache = CacheService.getScriptCache();
    let emailAgentsList = cache.get('emailAgentsList');
    let agentDemographicsList = cache.get('agentDemographicsList');
    let emailAgents = new Set();
    let agentDemographics = {};

    if (emailAgentsList && agentDemographicsList) {
      emailAgents = new Set(JSON.parse(emailAgentsList));
      agentDemographics = JSON.parse(agentDemographicsList);
    } else if (masterSheet && masterSheet.getLastRow() > 1) {
      const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 45).getValues();
      for (let i = 0; i < masterData.length; i++) {
        const ldap = masterData[i][0] ? masterData[i][0].toString().toLowerCase() : '';
        if (!ldap) continue;

        const channel = masterData[i][23] ? masterData[i][23].toString().toLowerCase() : '';
        if (channel === 'email') {
          emailAgents.add(ldap);
        }

        // Cache Supervisor (Col AB / 27) and Site (Col AS / 44)
        agentDemographics[ldap] = {
          supervisor: masterData[i][27] || 'Unknown',
          site: masterData[i][44] || 'Unknown'
        };
      }
      cache.put('emailAgentsList', JSON.stringify(Array.from(emailAgents)), 14400);
      cache.put('agentDemographicsList', JSON.stringify(agentDemographics), 14400);
    } else if (agentDemographicsList) {
      agentDemographics = JSON.parse(agentDemographicsList);
    }
"""

if search_str in content:
    content = content.replace(search_str, replace_str)
    with open(filepath, 'w') as f:
        f.write(content)
    print("Success 1")
else:
    print("Search string 1 not found")


search_str_2 = """
                agentsInInterval[ldap] = {
                  ldap: ldap,
                  sos: formattedSOS,
                  eos: formattedEOS,
                  site: site,
                  isOT: false,
                  casesLogged: 0,
                  regularCount: 0,
                  manualCount: 0,
                  reopenedCount: 0
                };
"""

replace_str_2 = """
                agentsInInterval[ldap] = {
                  ldap: ldap,
                  sos: formattedSOS,
                  eos: formattedEOS,
                  site: site,
                  supervisor: agentDemographics[ldap] ? agentDemographics[ldap].supervisor : 'Unknown',
                  isOT: false,
                  casesLogged: 0,
                  regularCount: 0,
                  manualCount: 0,
                  reopenedCount: 0
                };
"""

if search_str_2 in content:
    content = content.replace(search_str_2, replace_str_2)
    with open(filepath, 'w') as f:
        f.write(content)
    print("Success 2")
else:
    print("Search string 2 not found")


search_str_3 = """
          if (!agentsInInterval[ldap] && isOvertime) {
            agentsInInterval[ldap] = {
              ldap: ldap,
              sos: "OT",
              eos: "OT",
              site: site,
              isOT: true,
              otType: otType,
              casesLogged: 0,
              regularCount: 0,
              manualCount: 0,
              reopenedCount: 0,
              activityLogged: intervalActivity
            };
          }
"""

replace_str_3 = """
          if (!agentsInInterval[ldap] && isOvertime) {
            agentsInInterval[ldap] = {
              ldap: ldap,
              sos: "OT",
              eos: "OT",
              site: site,
              supervisor: agentDemographics[ldap] ? agentDemographics[ldap].supervisor : 'Unknown',
              isOT: true,
              otType: otType,
              casesLogged: 0,
              regularCount: 0,
              manualCount: 0,
              reopenedCount: 0,
              activityLogged: intervalActivity
            };
          }
"""

if search_str_3 in content:
    content = content.replace(search_str_3, replace_str_3)
    with open(filepath, 'w') as f:
        f.write(content)
    print("Success 3")
else:
    print("Search string 3 not found")
