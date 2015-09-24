using System.Threading.Tasks;

namespace MailSim.ProvidersREST
{
    public static class Helpers
    {
        public static TResult GetResult<TResult>(this Task<TResult> task)
        {
            return task
                .ConfigureAwait(false)
                .GetAwaiter()
                .GetResult();
        }

        public static void GetResult(this Task task)
        {
            task
                .ConfigureAwait(false)
                .GetAwaiter()
                .GetResult();
        }
    }
}
