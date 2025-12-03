using System;
using System.Windows.Forms;
using Sony.Vegas;

public class EntryPoint
{
    // This method is invoked by Vegas when the script runs.
    public void FromVegas(Vegas vegas)
    {
        // 1) Find a selected event
        TrackEvent selectedEvent = null;
        foreach (Track t in vegas.Project.Tracks)
        {
            foreach (TrackEvent e in t.Events)
            {
                if (e.Selected)
                {
                    selectedEvent = e;
                    break;
                }
            }
            if (selectedEvent != null) break;
        }

        if (selectedEvent == null)
        {
            MessageBox.Show("Select a video event first and then run this script.");
            return;
        }

        // 2) Choose which Video FX to add: change fxName to the installed FX name
        string fxName = "MyBezierFx"; // <<-- change this to your effect's displayed name
        VideoFxDescriptor descriptor = null;
        foreach (VideoFxDescriptor d in vegas.VideoFxDescriptors)
        {
            if (d.Name.Equals(fxName, StringComparison.OrdinalIgnoreCase))
            {
                descriptor = d;
                break;
            }
        }

        if (descriptor == null)
        {
            MessageBox.Show("Could not find Video FX: " + fxName + "\nInstalled descriptors count: " + vegas.VideoFxDescriptors.Count);
            return;
        }

        // 3) Add the FX to the event (slot 0)
        VideoEventFx fx = selectedEvent.AddVideoEffect(descriptor, 0);

        // 4) Find the numeric parameter you want to automate (change paramName as needed)
        string paramName = "Curve"; // <<-- change to the parameter name inside the FX (case-sensitive)
        VideoParameter targetParam = null;
        foreach (VideoParameter p in fx.Parameters)
        {
            if (p.Name == paramName)
            {
                targetParam = p;
                break;
            }
        }

        if (targetParam == null)
        {
            MessageBox.Show("Parameter not found: " + paramName + "\nAvailable parameters:\n" + String.Join(", ", Array.ConvertAll(fx.Parameters.ToArray(), x => x.Name)));
            return;
        }

        // 5) Create an automation envelope for that parameter (some versions expose Envelope directly)
        // NOTE: API names can vary by Vegas version. Many versions expose the Envelope via parameter.Automation or parameter.Envelope.
        // Try parameter.Automation or parameter.Envelope if one exists. This example uses a safe approach: try known names.
        AutomationParameter automation = null;
        try
        {
            // Many Vegas versions: VideoParameter.Automation
            automation = targetParam.Automation;
        }
        catch
        {
            try
            {
                // Some variants: VideoParameter.Envelope
                dynamic dyn = targetParam;
                automation = dyn.Envelope as AutomationParameter;
            }
            catch
            {
                automation = null;
            }
        }

        // If automation object is null, try to create an envelope through the parameter API (if available)
        if (automation == null)
        {
            try
            {
                // Some versions provide CreateAutomation() or CreateEnvelope()
                dynamic dynParam = targetParam;
                automation = dynParam.CreateAutomation();
            }
            catch
            {
                // Could not create automation programmatically — inform the user.
                MessageBox.Show("Could not access or create an automation envelope programmatically for parameter: " + paramName + "\nYou may need to add keyframes manually in the UI or inspect the parameter object via the scripting Object Browser for your Vegas version.");
                return;
            }
        }

        // 6) Clear existing points and add two key points: at event start -> 0.0, event end -> 1.0
        double t0 = selectedEvent.Start; // in ticks/time units used by envelope API
        double t1 = selectedEvent.End;

        // Clear existing points (API name may be Clear() or RemoveAll())
        try { automation.Clear(); } catch { try { automation.RemoveAll(); } catch { } }

        // Add points / keyframes. Method names and exact argument types differ by Vegas version.
        // Common pattern: automation.Points.Add(new AutomationPoint(time, value));
        // or automation.AddPoint(time, value)
        // We'll attempt multiple common ways (use whichever matches your API).
        try
        {
            // Try: AddPoint(time, value)
            automation.AddPoint(t0, 0.0);
            automation.AddPoint(t1, 1.0);
        }
        catch
        {
            try
            {
                // Try: Points.Add(new AutomationPoint(time, value))
                dynamic pts = automation.Points;
                Type apType = typeof(object); // placeholder
                // Construct point dynamically
                dynamic p0 = Activator.CreateInstance(pts.GetType().GetGenericArguments()[0], new object[] { t0, 0.0 });
                dynamic p1 = Activator.CreateInstance(pts.GetType().GetGenericArguments()[0], new object[] { t1, 1.0 });
                pts.Add(p0);
                pts.Add(p1);
            }
            catch
            {
                MessageBox.Show("Could not add automation points programmatically — API differences between Vegas versions prevented the script from calling AddPoint. Use the scripting Object Browser to find the correct method names for automation points on your version.");
                return;
            }
        }

        // 7) Set interpolation/tangents to Bezier/Smooth for those points (API varies).
        // Many versions expose per-point properties like Point.Interpolation or LeftTangent/RightTangent types.
        // This block attempts to set common properties and will silently continue if unavailable.
        try
        {
            var points = automation.Points;
            // set first point right tangent, and second point left tangent to Bezier/Smooth
            if (points.Count >= 2)
            {
                dynamic p0 = points[0];
                dynamic p1 = points[points.Count - 1];

                // Common property names:
                try { p0.RightTangentType = TangentType.Smooth; } catch { }
                try { p1.LeftTangentType = TangentType.Smooth; } catch { }

                // Some versions use Interpolation property:
                try { p0.Interpolation = InterpolationType.Smooth; } catch { }
                try { p1.Interpolation = InterpolationType.Smooth; } catch { }

                // If there are explicit handle/control point properties, you can set them to shape the curve:
                // p0.RightHandle = new PointF(...); p1.LeftHandle = new PointF(...);
            }
        }
        catch
        {
            // ignore: interpolation setting is best-effort; the host UI still allows you to edit to Bezier manually.
        }

        // 8) Refresh UI
        try { vegas.Project.Synchronize(); } catch { }
        MessageBox.Show("Added FX and automation (attempted). If automation points were created but interpolation is not Bezier, open the envelope in the UI and set tangents to Smooth/Bezier (or adjust via the Object Browser API names for your Vegas version).");
    }
}
