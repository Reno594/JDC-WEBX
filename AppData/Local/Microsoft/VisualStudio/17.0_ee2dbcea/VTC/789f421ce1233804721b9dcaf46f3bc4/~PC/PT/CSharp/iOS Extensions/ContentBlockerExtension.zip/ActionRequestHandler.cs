using Foundation;

namespace $safeprojectname$
{
	[Register ("ActionRequestHandler")]
	public class ActionRequestHandler : NSExtensionRequestHandling
	{
		public override void BeginRequestWithExtensionContext (NSExtensionContext context)
		{
			var attachment = new NSItemProvider (NSBundle.MainBundle.GetUrlForResource ("blockerList", "json"));

			var item = new NSExtensionItem {
				Attachments = new [] { attachment }
			};

			context.CompleteRequest (new[] { item }, null);
		}
	}
}


