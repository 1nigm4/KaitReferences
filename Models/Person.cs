using System;

namespace KaitReferences.Models
{
    class Person
    {
        public string EmailAddress { get; set; }
        public string LastName { get; set; }
        public string Name { get; set; }
        public string Patronymic { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
        public DateTime BirthDate { get; set; }
        public string Gender { get; set; }
        public Education Education { get; set; }
        public Reference Reference { get; set; }

        public override string ToString()
        {
            return $"{LastName} {Name} {Patronymic}";
        }
    }
}
