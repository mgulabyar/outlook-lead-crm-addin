// /* global Office */
// import React, { useState } from "react";
// import {
//   Box, Typography, Button, Paper, CircularProgress, 
//   Fade, Stack, LinearProgress, Alert, Chip, Card
// } from "@mui/material";
// import {
//   Hub, PersonSearch, SyncAlt,  AccountCircle,
//   Business, Work, ContactPhone, AutoGraph, 
//   VerifiedUserOutlined
// } from "@mui/icons-material";
// import axios from "axios";

// const App: React.FC = () => {
//   const [loading, setLoading] = useState(false);
//   const [lead, setLead] = useState<any>(null);
//   const [error, setError] = useState<string | null>(null);

// const extractLead = async () => {
//   setLoading(true);
//   setLead(null);
//   setError(null);

//   try {
//     // 1. Check if Office is initialized
//     if (typeof Office === "undefined" || !Office.context || !Office.context.mailbox) {
//       setError("Outlook context not detected. Please run this inside Outlook.");
//       setLoading(false);
//       return;
//     }

//     const item = Office.context.mailbox.item;

//     // 2. Fix: 'item' is possibly 'undefined' error
//     if (!item) {
//       setError("No email selected.");
//       setLoading(false);
//       return;
//     }

//     // 3. Optional chaining to safely get sender email
//     const senderEmail = item.from?.emailAddress || "unknown@test.com";

//     // 4. Proper typing for the callback result
//     item.body.getAsync(Office.CoercionType.Text, async (result: Office.AsyncResult<string>) => {
//       if (result.status === Office.AsyncResultStatus.Failed) {
//         setError("Failed to capture email content.");
//         setLoading(false);
//         return;
//       }

//       const emailBody = result.value || "";

//       try {
//         // Backend API Call
//         const res = await axios.post("http://127.0.0.1:8003/api/leads/extract-and-sync", {
//           sender: senderEmail,
//           body: emailBody,
//         });

//         if (res.data && res.data.data) {
//           setLead(res.data.data);
//         } else {
//           setError("Lead could not be extracted.");
//         }
//       } catch (err: any) {
//         setError("Server Unreachable: Ensure backend runs on port 8003.");
//       } finally {
//         setLoading(false);
//       }
//     });

//   } catch (e) {
//     console.error("Extraction Error:", e);
//     setError("An unexpected error occurred.");
//     setLoading(false);
//   }
// };



//   const getPriorityColor = (p: string) => {
//     const priority = p?.toLowerCase();
//     if (priority === "hot") return "#dc2626"; 
//     if (priority === "warm") return "#ea580c"; 
//     return "#4b5563"; 
//   };

//   return (
//     <Box sx={{ height: "100vh", bgcolor: "#f1f5f9", display: "flex", flexDirection: "column" }}>
      
//       {/* PROFESSIONAL ENTERPRISE HEADER */}
//       <Paper elevation={0} sx={{ p: 2, bgcolor: "#0f172a", color: "white", borderRadius: 0, textAlign: "left" }}>
//         <Stack direction="row" spacing={1.5} alignItems="center">
//           <Box sx={{ bgcolor: "#38bdf8", p: 0.5, borderRadius: 1, display: "flex" }}>
//             <Hub sx={{ color: "#0f172a", fontSize: 22 }} />
//           </Box>
//           <Box>
//             <Typography variant="subtitle1" fontWeight={900} sx={{ lineHeight: 1, letterSpacing: 0.5 }}>
//               SALESFLOW AI
//             </Typography>
//             <Typography variant="caption" sx={{ color: "#94a3b8", fontWeight: 700, textTransform: "uppercase", fontSize: "10px" }}>
//               Lead Intelligence Engine
//             </Typography>
//           </Box>
//         </Stack>
//       </Paper>

//       <Box sx={{ p: 2, flexGrow: 1, overflowY: "auto" }}>
//         {error && <Alert severity="error" sx={{ mb: 2, borderRadius: 2 }}>{error}</Alert>}

//         {/* WORKFLOW PIPELINE */}
//         <Typography variant="caption" fontWeight={800} color="textSecondary" sx={{ mb: 1.5, display: "block", textAlign: "left", letterSpacing: 1 }}>
//           CRM PIPELINE STATUS
//         </Typography>

//         <Stack spacing={1} sx={{ mb: 3 }}>
//           {[
//             { label: "AI Lead Mining", icon: <PersonSearch sx={{ fontSize: 18 }} /> },
//             { label: "MongoDB Cloud Sync", icon: <SyncAlt sx={{ fontSize: 18 }} /> },
//           ].map((tool, i) => (
//             <Paper key={i} variant="outlined" sx={{ p: 1.2, borderRadius: 2, display: "flex", alignItems: "center", gap: 2, bgcolor: "white", borderColor: "#e2e8f0" }}>
//               <Box sx={{ color: "#0284c7", display: "flex" }}>{tool.icon}</Box>
//               <Typography variant="body2" fontWeight={700} color="#334155" sx={{ fontSize: "12px" }}>{tool.label}</Typography>
//               <VerifiedUserOutlined sx={{ ml: "auto", fontSize: 14, color: "#10b981" }} />
//             </Paper>
//           ))}
//         </Stack>

//         <Button
//           fullWidth variant="contained" disableElevation
//           onClick={extractLead} disabled={loading}
//           sx={{ bgcolor: "#0284c7", py: 1.4, fontWeight: 800, borderRadius: 2, mb: 3, "&:hover": { bgcolor: "#0369a1" }, textTransform: "none", fontSize: "14px" }}
//         >
//           {loading ? <CircularProgress size={22} sx={{ color: "white" }} /> : "Capture & Sync Lead"}
//         </Button>

//         {lead && (
//           <Fade in={true}>
//             <Box sx={{ textAlign: "left" }}>
              
//               {/* ANALYTICS CARD */}
//               <Card variant="outlined" sx={{ p: 2, mb: 3, borderRadius: 2, borderLeft: `6px solid ${getPriorityColor(lead.priority)}`, bgcolor: "white" }}>
//                 <Stack direction="row" justifyContent="space-between" alignItems="center" sx={{ mb: 1 }}>
//                   <Typography variant="caption" fontWeight={900} color="textSecondary">CONVERSION SCORE</Typography>
//                   <Chip label={lead.priority.toUpperCase()} size="small" sx={{ fontWeight: 900, fontSize: "10px", height: 20, bgcolor: "#f8fafc", color: getPriorityColor(lead.priority), border: `1px solid ${getPriorityColor(lead.priority)}` }} />
//                 </Stack>
//                 <Stack direction="row" alignItems="baseline" spacing={1}>
//                   <Typography variant="h3" fontWeight={900} sx={{ color: "#0f172a" }}>{lead.lead_score}</Typography>
//                   <Typography variant="body2" color="textSecondary" fontWeight={700}>/ 100</Typography>
//                 </Stack>
//                 <LinearProgress variant="determinate" value={lead.lead_score} sx={{ height: 8, borderRadius: 5, mt: 1.5, bgcolor: "#f1f5f9", "& .MuiLinearProgress-bar": { bgcolor: getPriorityColor(lead.priority) } }} />
//               </Card>

//               {/* DATA ATTRIBUTES */}
//               <Typography variant="caption" fontWeight={800} color="textSecondary" sx={{ mb: 1.5, display: "block", letterSpacing: 1 }}>
//                 EXTRACTED CRM DATA
//               </Typography>

//               <Paper variant="outlined" sx={{ p: 2, borderRadius: 2, bgcolor: "white", mb: 3 }}>
//                 <Stack spacing={2.5}>
//                   {[
//                     { label: "FULL NAME", val: lead.name, icon: <AccountCircle /> },
//                     { label: "ORGANIZATION", val: lead.company, icon: <Business /> },
//                     { label: "JOB TITLE", val: lead.designation, icon: <Work /> },
//                     { label: "CONTACT", val: lead.phone, icon: <ContactPhone /> },
//                   ].map((item, i) => (
//                     <Stack key={i} direction="row" spacing={2} alignItems="center">
//                       <Box sx={{ color: "#64748b", display: "flex" }}>{item.icon}</Box>
//                       <Box>
//                         <Typography variant="caption" sx={{ fontSize: "10px", fontWeight: 800, color: "#94a3b8", display: "block", lineHeight: 1 }}>
//                           {item.label}
//                         </Typography>
//                         <Typography variant="body2" fontWeight={700} color="#1e293b">
//                           {item.val || "Not specified"}
//                         </Typography>
//                       </Box>
//                     </Stack>
//                   ))}
//                 </Stack>
//               </Paper>

//               {/* AI INSIGHT */}
//               <Box sx={{ p: 2, bgcolor: "#f0f9ff", borderRadius: 2, border: "1px solid #bae6fd" }}>
//                 <Stack direction="row" spacing={1} alignItems="center" sx={{ mb: 1 }}>
//                   <AutoGraph sx={{ fontSize: 18, color: "#0369a1" }} />
//                   <Typography variant="caption" fontWeight={900} color="#0369a1" sx={{ letterSpacing: 0.5 }}>STRATEGIC INTENT</Typography>
//                 </Stack>
//                 <Typography variant="body2" sx={{ color: "#075985", fontStyle: "italic", lineHeight: 1.5, fontWeight: 500 }}>
//                   "{lead.intent_summary}"
//                 </Typography>
//               </Box>

//             </Box>
//           </Fade>
//         )}
//       </Box>





import React, { useState } from "react";
import {
  Box, Typography, Button, Paper, CircularProgress, 
  Fade, Stack, LinearProgress, Alert} from "@mui/material";
import {
  AccountCircle,
  Business, Work, ContactPhone, AutoGraph, 
  ErrorOutline
} from "@mui/icons-material";
import axios from "axios";

const App: React.FC = () => {
  const [loading, setLoading] = useState(false);
  const [lead, setLead] = useState<any>(null);
  const [error, setError] = useState<string | null>(null);

  // --- LOGIC: MAINTAINED AS PER YOUR ORIGINAL CODE ---
  const extractLead = async () => {
    setLoading(true);
    setLead(null);
    setError(null);

    try {
      if (typeof Office === "undefined" || !Office.context || !Office.context.mailbox) {
        setError("Outlook context not detected.");
        setLoading(false);
        return;
      }
      const item = Office.context.mailbox.item;
      if (!item) {
        setError("No email selected.");
        setLoading(false);
        return;
      }
      const senderEmail = item.from?.emailAddress || "unknown@test.com";
      item.body.getAsync(Office.CoercionType.Text, async (result: Office.AsyncResult<string>) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          setError("Failed to capture content.");
          setLoading(false);
          return;
        }
        const emailBody = result.value || "";
        try {
          const res = await axios.post("http://127.0.0.1:8003/api/leads/extract-and-sync", {
            sender: senderEmail,
            body: emailBody,
          });
          if (res.data && res.data.data) {
            setLead(res.data.data);
          } else {
            setError("Lead extraction failed.");
          }
        } catch (err: any) {
          setError("Server Unreachable: Port 8003.");
        } finally {
          setLoading(false);
        }
      });
    } catch (e) {
      setError("An unexpected error occurred.");
      setLoading(false);
    }
  };

  const getPriorityColor = (p: string) => {
    const priority = p?.toLowerCase();
    if (priority === "hot") return "#B91C1C"; // Corporate Red
    if (priority === "warm") return "#C49A6C"; // Bronze/Gold
    return "#475569"; // Slate
  };

  return (
    <Box sx={{ 
      height: "100vh", bgcolor: "#F8FAFC", display: "flex", flexDirection: "column",
      fontFamily: "'Roboto', 'Helvetica', sans-serif" 
    }}>
      
      {/* 1. STICKY PROFESSIONAL HEADER (No Logo) */}
      <Box sx={{ 
        position: "sticky", top: 0, zIndex: 1000, 
        bgcolor: "#1B1B42", color: "white", p: 2, 
        textAlign: "center", borderBottom: "3px solid #38BDF8"
      }}>
        <Typography variant="h6" sx={{ fontWeight: 800, letterSpacing: "1.5px", fontSize: "18px" }}>
          SALESFLOW AI
        </Typography>
        <Typography sx={{ fontSize: "9px", fontWeight: 500, color: "#38BDF8", textTransform: "uppercase", letterSpacing: "2px", mt: 0.5 }}>
          Enterprise Lead Intelligence
        </Typography>
      </Box>

      <Box sx={{ p: 2, flexGrow: 1, overflowY: "auto" }}>
        {error && <Alert severity="error" icon={<ErrorOutline />} sx={{ mb: 2, fontSize: "11px", borderRadius: "8px" }}>{error}</Alert>}

        {/* 2. SYSTEM CONTEXT BOXES (Replacement for tool features) */}
        <Box sx={{ mb: 2.5 }}>
           <Typography variant="body2" sx={{ fontSize: "10px", color: "#94A3B8", textTransform: "uppercase", fontWeight: 700, mb: 1, ml: 0.5 }}>
             Automation Workflow
           </Typography>
           
           <Paper variant="outlined" sx={{ p: 1.5, mb: 1, borderRadius: "10px", borderLeft: "4px solid #1B1B42" }}>
              <Typography sx={{ fontSize: "11.5px", fontWeight: 600, color: "#1E293B" }}>Neural Data Extraction</Typography>
              <Typography sx={{ fontSize: "10.5px", color: "#64748B", mt: 0.5 }}>Utilize AI to parse unstructured email bodies into high-fidelity CRM profiles.</Typography>
           </Paper>

           <Paper variant="outlined" sx={{ p: 1.5, borderRadius: "10px", borderLeft: "4px solid #38BDF8" }}>
              <Typography sx={{ fontSize: "11.5px", fontWeight: 600, color: "#1E293B" }}>Cloud Ledger Sync</Typography>
              <Typography sx={{ fontSize: "10.5px", color: "#64748B", mt: 0.5 }}>Direct synchronization with centralized databases and enterprise CRM nodes.</Typography>
           </Paper>
        </Box>

        {/* 3. COMPACT CTA BUTTON */}
        <Button
          fullWidth variant="contained" disableElevation
          onClick={extractLead} disabled={loading}
          sx={{ 
            bgcolor: "#1B1B42", py: 1.2, fontWeight: 800, borderRadius: "6px", mb: 3, 
            fontSize: "12px", textTransform: "none",
            "&:hover": { bgcolor: "#2A2A5A" },
            "&.Mui-disabled": { bgcolor: "#1B1B42", opacity: 0.7, color: "white" }
          }}
        >
          {loading ? <CircularProgress size={18} color="inherit" /> : "Initiate Lead Capture"}
        </Button>

        {/* 4. RESULT DATA SECTION */}
        {lead && (
          <Fade in={true}>
            <Box>
              {/* SCORE BOX */}
              <Box sx={{ 
                p: 2, mb: 2.5, borderRadius: "12px", border: "1px solid #E2E8F0", 
                bgcolor: "white", borderLeft: `6px solid ${getPriorityColor(lead.priority)}`
              }}>
                <Box sx={{ display: "flex", justifyContent: "space-between", mb: 1 }}>
                  <Typography sx={{ fontSize: "10px", fontWeight: 800, color: "#64748B" }}>CONVERSION RISK</Typography>
                  <Typography sx={{ fontSize: "10px", fontWeight: 800, color: getPriorityColor(lead.priority) }}>
                    {lead.priority.toUpperCase()}
                  </Typography>
                </Box>
                <Box sx={{ display: "flex", alignItems: "baseline" }}>
                  <Typography sx={{ fontSize: "32px", fontWeight: 800, color: "#1E293B" }}>{lead.lead_score}</Typography>
                  <Typography sx={{ fontSize: "12px", ml: 0.5, color: "#94A3B8", fontWeight: 700 }}>/ 100</Typography>
                </Box>
                <LinearProgress variant="determinate" value={lead.lead_score} sx={{ mt: 1.5, height: 6, borderRadius: 3, bgcolor: "#F1F5F9", "& .MuiLinearProgress-bar": { bgcolor: getPriorityColor(lead.priority) } }} />
              </Box>

              {/* DATA ATTRIBUTES BOX */}
              <Box sx={{ p: 2, bgcolor: "white", border: "1px solid #E2E8F0", borderRadius: "12px" }}>
                <Typography sx={{ fontSize: "10px", fontWeight: 800, color: "#94A3B8", mb: 2, textTransform: "uppercase" }}>Extracted Attributes</Typography>
                <Stack spacing={2}>
                  {[
                    { label: "FULL NAME", val: lead.name, icon: <AccountCircle sx={{fontSize: 16}}/> },
                    { label: "ORGANIZATION", val: lead.company, icon: <Business sx={{fontSize: 16}}/> },
                    { label: "DESIGNATION", val: lead.designation, icon: <Work sx={{fontSize: 16}}/> },
                    { label: "CONTACT LINE", val: lead.phone, icon: <ContactPhone sx={{fontSize: 16}}/> },
                  ].map((item, i) => (
                    <Box key={i} sx={{ display: "flex", alignItems: "center" }}>
                      <Box sx={{ color: "#1B1B42", display: "flex", mr: 1.5, opacity: 0.8 }}>{item.icon}</Box>
                      <Box>
                        <Typography sx={{ fontSize: "8.5px", fontWeight: 800, color: "#94A3B8", display: "block" }}>{item.label}</Typography>
                        <Typography sx={{ fontSize: "12px", fontWeight: 700, color: "#1E293B" }}>{item.val || "N/A"}</Typography>
                      </Box>
                    </Box>
                  ))}
                </Stack>
              </Box>

              {/* AI SUMMARY BOX */}
              <Box sx={{ p: 1.5, bgcolor: "#F0F9FF", border: "1px solid #E0F2FE", borderRadius: "8px", mt: 2 }}>
                <Box sx={{ display: "flex", alignItems: "center", mb: 0.5 }}>
                  <AutoGraph sx={{ fontSize: 14, color: "#0284C7", mr: 1 }} />
                  <Typography sx={{ fontSize: "10px", fontWeight: 800, color: "#0284C7" }}>AI STRATEGY INSIGHT</Typography>
                </Box>
                <Typography sx={{ fontSize: "11px", color: "#0C4A6E", fontStyle: "italic", lineHeight: 1.5, fontWeight: 500 }}>
                  "{lead.intent_summary}"
                </Typography>
              </Box>
            </Box>
          </Fade>
        )}
      </Box>

      {/* 5. SYSTEM FOOTER */}
      <Box sx={{ p: 1.2, textAlign: "center", bgcolor: "#FFFFFF", borderTop: "1px solid #E2E8F0" }}>
        <Typography sx={{ fontSize: "9px", fontWeight: 800, color: "#CBD5E1", letterSpacing: "1px" }}>
          PRODUCTION NODE v4.5 | SECURE PIPELINE
        </Typography>
      </Box>
    </Box>
  );
};

export default App;