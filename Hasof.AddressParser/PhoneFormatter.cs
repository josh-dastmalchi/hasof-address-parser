namespace Hasof.AddressParser
{
    public static class PhoneFormatter
    {
        public static string Format(string unformatted)
        {
            if (string.IsNullOrWhiteSpace(unformatted))
            {
                return string.Empty;
            }

            var withoutPunctation = unformatted.Replace("-", string.Empty).Replace(" ", string.Empty);
            if (withoutPunctation.Length == 7)
            {
                return withoutPunctation.Substring(0, 3) + "-" + withoutPunctation.Substring(3, 4);
            }
            if (withoutPunctation.Length == 10)
            {
                return withoutPunctation.Substring(0, 3) + "-" + withoutPunctation.Substring(3, 3) + "-" + withoutPunctation.Substring(6);
            }
            
            return unformatted;
        }
    }
}
