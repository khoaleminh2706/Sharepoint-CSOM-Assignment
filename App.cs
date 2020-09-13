﻿using System;
using CreateSPSite.Factories;
using CreateSPSite.Models;
using CreateSPSite.Services;

namespace CreateSPSite
{
    public class App
    {
        private bool _over;
        private ConsoleKeyInfo _key;
        private string _loginName;
        private string _password;
        private string _siteUrl;
        private SPClientContextProvider _provider;
        private ContentTypeFactory _contentTypeFactory;

        public App(bool over)
        {
            _over = over;
        }

        public void Run()
        {
            // TODO: allow changing url
            Console.WriteLine("Welcome...");
            Console.WriteLine("Vui lòng điền thông tin đăng nhập:");
            Console.Write("Login Name: ");
            _loginName = Console.ReadLine();
            Console.Write("Password: ");
            _password = Console.ReadLine();
            Console.Write("Site Url: ");
            _siteUrl = Console.ReadLine();

            _provider = new SPClientContextProvider(_loginName, _password, _siteUrl);
            _contentTypeFactory = new ContentTypeFactory(_provider.Create());

            while (!_over)
            {
                Update(); 
            }
            Console.WriteLine("Exiting...");
        }

        private void Update()
        {
            Console.WriteLine("Please select 1 action");
            Console.WriteLine("[1] Create Employees list");
            Console.WriteLine("[2] Create Project list");
            Console.WriteLine("[3] Create Project Document list");
            Console.WriteLine("[4] Create Site");
            Console.WriteLine("[5] Change Site Url");
            Console.WriteLine("[Esc or Ctrl-C] Exit");
            _key = Console.ReadKey();

            // xuống 1 dòng
            Console.WriteLine();
            HandleKey(_key);
        }

        public void HandleKey(ConsoleKeyInfo key)
        {
            switch (key.Key)
            {
                case ConsoleKey.D1:
                    try
                    {
                        _contentTypeFactory.GetContentType(Constants.ConteType.Employee);
                    }
                    catch(Exception ex)
                    {
                        Console.WriteLine("Chương trình bị lỗi");
                        Console.WriteLine(ex.Message);
                    }
                    break;
                case ConsoleKey.D2:
                    //SharepointService.CreateProjectList1();
                    break;
                case ConsoleKey.D3:
                    Console.WriteLine("You press 3");
                    // TODO: Delete Project Documents
                    break;
                case ConsoleKey.D4:
                    // TODO: Create Site and Sub site
                    break;
                case ConsoleKey.D5:
                    // TODO: allow changing site url
                    break;
                case ConsoleKey.Escape:
                    _over = true;
                    break;
                default:
                    return;
            }
        }

        private void HandleCreateList(string listName)
        {
            var context = _provider.Create();
        }
    }
}
